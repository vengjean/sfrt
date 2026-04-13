# -*- coding: utf-8 -*-
"""
Updated on May 9th, 2024
@author: Veng Jean Heng
# Main script: SFRT
"""

from connect import *

import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

from System.Drawing import Point, Color
from System.Windows.Forms import Application, Button, Form, Label, ComboBox, TextBox, ComboBoxStyle
from System.Windows import MessageBox
from System.ComponentModel import BackgroundWorker

from string import ascii_uppercase
import csv
import os


def find_default_gtv_name(total_roi_list):
    gtv_name = "GTV"
    if gtv_name not in total_roi_list:
        for roi in total_roi_list:
            if "GTV" in roi:
                gtv_name = roi
                break

        if gtv_name not in total_roi_list:
            raise Exception("GTV not found")
    return gtv_name


class ClinicalGoals(Form):
    def __init__(self, plan, exam, spheresize=1.5, presc_ld=2000):
        self.case = get_current("Case")
        self.plan = plan

        self.roi_list = []
        for roi in self.case.PatientModel.RegionsOfInterest:
            self.roi_list.append(roi.Name)

        self.oar_filename = 'SFRT_ClinicalGoals_OAR.csv'
        self.filename = 'SFRT_ClinicalGoals'
        if spheresize == 1.0:
            self.filename = self.filename + '_10mm'
        else:
            self.filename = self.filename + '_15mm'

        if presc_ld == 1000:
            self.filename = self.filename + '_10Gy.csv'
        else:
            self.filename = self.filename + '_20Gy.csv'

        if not os.path.isfile(self.filename):
            script_dir = os.path.dirname(os.path.abspath(__file__))
            self.filename = os.path.join(script_dir, self.filename)
            self.oar_filename = os.path.join(script_dir, self.oar_filename)

        self.add_clinical_goals(self.filename, self.oar_filename, exam)

    def open_csv(self, filename):
        goals = []
        with open(filename, 'r') as f:
            reader = csv.reader(f)
            next(reader)  # skip header
            for row in reader:
                goals.append(row)
        return goals

    def add_clinical_goal(self, goal, exam):
        roi_names = goal[0].split(';')
        for roi_name in roi_names:
            if roi_name in self.roi_list:
                # Check if the roi has a contour
                if self.case.PatientModel.StructureSets[exam].RoiGeometries[roi_name].HasContours():
                    self.plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(
                        RoiName=roi_name,
                        GoalCriteria=goal[1],
                        GoalType=goal[2],
                        AcceptanceLevel=float(goal[3]),
                        ParameterValue=float(goal[4]),
                        IsComparativeGoal=False)
                    break

    def add_clinical_goals(self, filename, oar, exam):
        with CompositeAction('Add Clinical Goals'):
            for goal in self.open_csv(filename):
                self.add_clinical_goal(goal, exam)
            for goal in self.open_csv(oar):
                self.add_clinical_goal(goal, exam)


class SFRTForm(Form):
    def __init__(self):

        self.case = get_current("Case")

        self.plan_prefix = ""
        found_suffix = []
        for tp in self.case.TreatmentPlans:
            tp_suffix = tp.Name.split('_')[-1]
            found_suffix.append(tp_suffix)
        for idx, alph in enumerate(ascii_uppercase):
            if alph not in found_suffix:
                temp_suffix_idx = idx
                break
        self.plan_suffix = ""

        # Select planning/ref CT:
        refCTs = []
        for ixt, listCT in enumerate(self.case.Examinations):
            if listCT.EquipmentInfo.Modality == 'CT':
                refCTs.append(self.case.Examinations[ixt].Name)

        self.Text = 'SFRT Planning Interface'
        self.Height = 500
        self.Width = 364
        self.ForeColor = Color.Black
        self.BackColor = Color.LightGray

        # Description label at the top of the GUI
        self.label1 = Label()
        self.label1.Text = "Generate a new RT Plan for SFRT"
        self.label1.Location = Point(10, 10)
        self.label1.Height = 20
        self.label1.Width = 400

        # Plan name field form
        self.label2 = Label()
        self.label2.Text = "Plan name prefix"
        self.label2.Location = Point(10, 40)
        self.label2.Height = 20
        self.label2.Width = 150

        self.tb_plan_name = TextBox()
        self.tb_plan_name.Parent = self
        self.tb_plan_name.Location = Point(10, 60)
        self.tb_plan_name.Width = 150
        self.tb_plan_name.Text = self.plan_prefix
        self.tb_plan_name.TextChanged += self.PlanLetterOnChanged

        # Plan letter field form
        self.label3 = Label()
        self.label3.Text = "MOSAIQ letter suffix"
        self.label3.Location = Point(200, 40)
        self.label3.Height = 20
        self.label3.Width = 200

        letters = [s for s in ascii_uppercase]
        # convert to tuple
        letters = tuple(letters)
        self.cb_plan_letter = ComboBox()
        self.cb_plan_letter.Parent = self
        self.cb_plan_letter.Location = Point(200, 60)
        self.cb_plan_letter.Items.AddRange(letters)
        self.cb_plan_letter.SelectedIndex = temp_suffix_idx
        self.cb_plan_letter.SelectionChangeCommitted += self.PlanLetterOnChanged
        self.plan_suffix = self.cb_plan_letter.SelectedItem
        self.cb_plan_letter.DropDownStyle = ComboBoxStyle.DropDownList

        self.label4 = Label()
        self.label4.Text = "Full plan name: "
        self.label4.Location = Point(10, 80)
        self.label4.Height = 20
        self.label4.Width = 400
        self.PlanLetterOnChanged(None, None)

        # Label: Prescription dose
        self.label5 = Label()
        self.label5.Text = "Prescription dose"
        self.label5.Location = Point(10, 110)
        self.label5.Height = 20
        self.label5.Width = 180

        # Combobox: Prescription dose
        self.cb_presc = ComboBox()
        self.cb_presc.Parent = self
        self.cb_presc.Location = Point(10, 130)
        self.cb_presc.Items.AddRange(
            ("HD 66.7 Gy / LD 20 Gy", "HD 66.7 Gy / LD 10 Gy"))
        self.cb_presc.SelectionChangeCommitted += self.PrescriptionOnChanged
        self.cb_presc.SelectedIndex = 0
        self.cb_presc.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_presc.Width = 160
        self.PrescriptionOnChanged(None, None)

        # Label: Sphere size
        self.label6 = Label()
        self.label6.Text = "Sphere diameter"
        self.label6.Location = Point(200, 110)
        self.label6.Height = 20
        self.label6.Width = 100

        # Combobox: Sphere size
        self.cb_sphere = ComboBox()
        self.cb_sphere.Parent = self
        self.cb_sphere.Location = Point(200, 130)
        self.cb_sphere.Items.AddRange(
            ("1.0 cm", "1.5 cm")
        )
        self.cb_sphere.SelectedIndex = 1
        self.cb_sphere.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_sphere.SelectionChangeCommitted += self.SphereOnChanged
        self.SphereOnChanged(None, None)

        # Label: Number of beams
        self.label7 = Label()
        self.label7.Text = "Number of beams"
        self.label7.Location = Point(10, 170)
        self.label7.Height = 20
        self.label7.Width = 180

        # Combobox: Number of beams
        self.cb_nb = ComboBox()
        self.cb_nb.Parent = self
        self.cb_nb.Location = Point(10, 190)
        self.cb_nb.Items.AddRange(("3", "4", "5"))
        self.cb_nb.SelectionChangeCommitted += self.NBOnChanged
        self.cb_nb.SelectedIndex = 2
        self.cb_nb.DropDownStyle = ComboBoxStyle.DropDownList
        self.NBOnChanged(None, None)

        # Label: Gantry angles
        self.label8 = Label()
        self.label8.Text = "Gantry angles"
        self.label8.Location = Point(200, 170)
        self.label8.Height = 20
        self.label8.Width = 100

        # Textbox: Gantry angles 1
        self.tb_ga = TextBox()
        self.tb_ga.Parent = self
        self.tb_ga.Location = Point(200, 190)
        self.tb_ga.Width = 40
        self.tb_ga.Text = "181"

        # Textbox: Gantry angles 2
        self.tb_ga2 = TextBox()
        self.tb_ga2.Parent = self
        self.tb_ga2.Location = Point(260, 190)
        self.tb_ga2.Width = 40
        self.tb_ga2.Text = "179"

        # Label: Linac type
        self.label9 = Label()
        self.label9.Text = "Linac"
        self.label9.Location = Point(10, 230)
        self.label9.Height = 20
        self.label9.Width = 150

        # Combobox: Linac type
        self.cb_lt = ComboBox()
        self.cb_lt.Parent = self
        self.cb_lt.Location = Point(10, 250)
        self.cb_lt.Items.AddRange(
            ("TrueBeam", "TrueBeamFFF"))
        self.cb_lt.SelectionChangeCommitted += self.LTOnChanged
        self.cb_lt.SelectedIndex = 0
        self.cb_lt.DropDownStyle = ComboBoxStyle.DropDownList
        self.LTOnChanged(None, None)

        # Label: Energy
        self.label10 = Label()
        self.label10.Text = "Energy (MV)"
        self.label10.Location = Point(200, 230)
        self.label10.Height = 20
        self.label10.Width = 400

        # Combobox: Energy
        self.cb_energy = ComboBox()
        self.cb_energy.Parent = self
        self.cb_energy.Location = Point(200, 250)
        self.cb_energy.Items.AddRange(
            ("6", "10"))
        self.cb_energy.SelectionChangeCommitted += self.EnergyOnChanged
        self.cb_energy.SelectedIndex = 0
        self.cb_energy.DropDownStyle = ComboBoxStyle.DropDownList
        self.EnergyOnChanged(None, None)

        # Label: Planning CT
        self.label11 = Label()
        self.label11.Text = "Planning CT"
        self.label11.Location = Point(10, 280)
        self.label11.Height = 20
        self.label11.Width = 180

        # Combobox: Planning CT
        self.cb_ct = ComboBox()
        self.cb_ct.Parent = self
        self.cb_ct.Location = Point(10, 300)
        for i in refCTs:
            self.cb_ct.Items.Add(i)
        self.cb_ct.SelectionChangeCommitted += self.CTOnChanged
        self.cb_ct.DropDownStyle = ComboBoxStyle.DropDownList
        self.ct = 'None'

        # Label: External contour
        self.label12 = Label()
        self.label12.Text = "External contour"
        self.label12.Location = Point(200, 280)
        self.label12.Height = 20
        self.label12.Width = 150

        # ComboBox: External contour
        self.cb_external = ComboBox()
        self.cb_external.Parent = self
        self.cb_external.Location = Point(200, 300)
        self.cb_external.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_external.SelectionChangeCommitted += self.ExternalOnChanged
        self.external_name = 'External'

        # Label: Isocenter
        self.label13 = Label()
        self.label13.Text = "Isocenter"
        self.label13.Location = Point(10, 330)
        self.label13.Height = 20
        self.label13.Width = 200

        # ComboBox: Isocenter
        self.cb_isocenter = ComboBox()
        self.cb_isocenter.Parent = self
        self.cb_isocenter.Location = Point(10, 350)

        self.isocenter = 'None'
        self.cb_isocenter.SelectionChangeCommitted += self.ISOOnChanged
        self.cb_isocenter.DropDownStyle = ComboBoxStyle.DropDownList

        # Label: Isocenter warning message
        self.label14 = Label()
        self.label14.Text = "The isocenter must already be created in the selected CT dataset"
        self.label14.Location = Point(150, 350)
        self.label14.Width = 200
        self.label14.Height = 40

        # Label: Message
        self.label15 = Label()
        self.label15.Text = " "
        self.label15.Location = Point(10, 390)
        self.label15.Height = 20
        self.label15.Width = 400

        # Button: Generate new plan
        self.button1 = Button()
        self.button1.Text = "New Plan"
        self.button1.Location = Point(10, 430)
        self.button1.Click += self.buttonPressed
        self.button1.Enabled = False

        # Button: Force overwrite contour
        self.button2 = Button()
        self.button2.Text = "Overwrite contours"
        self.button2.Location = Point(200, 430)
        self.button2.Click += self.overwritePressed
        self.button2.Width = 120
        self.button2.Enabled = False

        self.Controls.Add(self.label1)
        self.Controls.Add(self.label2)
        self.Controls.Add(self.label3)
        self.Controls.Add(self.label4)
        self.Controls.Add(self.label5)
        self.Controls.Add(self.label6)
        self.Controls.Add(self.label7)
        self.Controls.Add(self.label8)
        self.Controls.Add(self.label9)
        self.Controls.Add(self.label10)
        self.Controls.Add(self.label11)
        self.Controls.Add(self.label12)
        self.Controls.Add(self.label13)
        self.Controls.Add(self.tb_plan_name)
        self.Controls.Add(self.label14)
        self.Controls.Add(self.label15)
        self.Controls.Add(self.button1)
        self.Controls.Add(self.button2)
        self.CenterToScreen()

    def ISOOnChanged(self, sender, event):
        isocenter_name = self.cb_isocenter.SelectedItem
        self.isocenter = self.case.PatientModel.StructureSets[
            self.ct].PoiGeometries[isocenter_name]

        if isocenter_name.upper() == "ISO_X":
            self.label14.Text = "Make sure that ISO_X has been correctly positioned beforehand"
            self.label14.ForeColor = Color.DarkRed
        else:
            self.label14.Text = " "
            self.label14.ForeColor = Color.Black

    def NBOnChanged(self, sender, event):
        self.nb = self.cb_nb.SelectedItem
        self.nob = int(self.nb)

    def LTOnChanged(self, sender, event):
        self.lt = self.cb_lt.SelectedItem

    def EnergyOnChanged(self, sender, event):
        self.energy = self.cb_energy.SelectedItem

    def ExternalOnChanged(self, sender, event):
        self.external_name = self.cb_external.SelectedItem

    def CTOnChanged(self, sender, event):
        self.ct = self.cb_ct.SelectedItem

        self.exam = self.case.Examinations[self.ct]
        self.button2.Enabled = True
        self.button1.Enabled = True

        self.cb_isocenter.Items.Clear()
        for poi in self.case.PatientModel.StructureSets[self.ct].PoiGeometries:
            if poi.OfPoi.Type == 'Isocenter':
                if poi.Point is not None:
                    self.cb_isocenter.Items.Add(poi.OfPoi.Name)
                    self.cb_isocenter.SelectedIndex = 0
                    self.ISOOnChanged(None, None)

        # Populate external contour dropdown
        self.cb_external.Items.Clear()
        external_names = []
        for roi in self.case.PatientModel.StructureSets[self.ct].RoiGeometries:
            if roi.OfRoi.Type == 'External' and roi.HasContours():
                external_names.append(roi.OfRoi.Name)
        for name in external_names:
            self.cb_external.Items.Add(name)
        # Default selection: "External" > contains "external" > contains "body" > first
        default_idx = 0
        for idx, name in enumerate(external_names):
            if name == "External":
                default_idx = idx
                break
            elif "external" in name.lower() and default_idx == 0:
                default_idx = idx
            elif "body" in name.lower() and default_idx == 0:
                default_idx = idx
        if external_names:
            self.cb_external.SelectedIndex = default_idx
            self.external_name = external_names[default_idx]

    def PrescriptionOnChanged(self, sender, event):
        if self.cb_presc.SelectedItem == "HD 66.7 Gy / LD 20 Gy":
            self.presc_hd = 6670
            self.presc_ld = 2000
        elif self.cb_presc.SelectedItem == "HD 66.7 Gy / LD 10 Gy":
            self.presc_hd = 6670
            self.presc_ld = 1000

    def SphereOnChanged(self, sender, event):
        if self.cb_sphere.SelectedItem == "1.0 cm":
            self.sphere_size = 1.0
        elif self.cb_sphere.SelectedItem == "1.5 cm":
            self.sphere_size = 1.5
        else:
            raise Exception("Sphere size not found")

    def PlanLetterOnChanged(self, sender, event):
        self.plan_suffix = self.cb_plan_letter.SelectedItem
        self.plan_prefix = self.tb_plan_name.Text
        if len(self.plan_prefix) > 16:
            self.plan_prefix = self.plan_prefix[:16]
            self.tb_plan_name.Text = self.plan_prefix
        self.plan_name = self.plan_prefix + '_' + self.plan_suffix
        self.label4.Text = 'Full plan name: ' + self.plan_name
        self.label4.ForeColor = Color.DarkRed

    def overwritePressed(self, sender, args):
        self.label15.Text = "Please wait for the contour overwrite to be completed..."
        self.label15.ForeColor = Color.DarkOrange
        self.button1.Enabled = False
        self.button2.Enabled = False

        self.worker = BackgroundWorker()
        self.worker.DoWork += self.overwrite_opt_structures
        self.worker.RunWorkerCompleted += self.overwrite_completed
        self.worker.RunWorkerAsync()

    def overwrite_opt_structures(self, sender, e):
        self.create_opt_structures(force_delete=True)

    def overwrite_completed(self, sender, e):
        self.worker.Dispose()
        if e.Error:
            MessageBox.Show(str(e.Error))
        else:
            self.label15.Text = "Optimization structures were overwritten successfully!"
            self.label15.ForeColor = Color.DarkGreen
            self.button1.Enabled = True
            self.button2.Enabled = True

    def buttonPressed(self, sender, args):
        self.label15.Text = "Please wait for the plan to be generated..."
        self.label15.ForeColor = Color.DarkOrange
        self.button1.Enabled = False
        self.button2.Enabled = False

        self.GA_start = int(self.tb_ga.Text)
        self.GA_stop = int(self.tb_ga2.Text)

        self.worker = BackgroundWorker()
        self.worker.DoWork += self.start_plancreation
        self.worker.RunWorkerCompleted += self.worker_completed
        self.worker.RunWorkerAsync()

    def worker_completed(self, sender, e):
        self.worker.Dispose()
        if e.Error:
            MessageBox.Show(str(e.Error))
        else:
            MessageBox.Show('New SFRT plan named ' + self.tb_plan_name.Text +
                            ' was generated successfully! \n Select the plan in the Plan List to view the plan.')
            self.window.Close()

    def start_plancreation(self, sender, e):

        # Check plan_prefix is less than 16 chars
        if len(self.plan_prefix) > 16:
            MessageBox.Show('Plan name should be less than 16 characters')
            self.button1.Enabled = True
            return
        self.nof = 5

        # Clinical Goals: background info for planning CT set
        self.exam = self.case.Examinations[self.ct]
        self.exam.SetPrimary()

        self.plan, self.beam_set = self.add_new_plan(
            self.case, self.ct, self.nob, self.presc_hd, self.nof, self.lt, self.plan_name)
        # Set current plan

        print('Creating optimization structures')
        self.label15.Text = "Creating optimization structures..."

        self.create_opt_structures()

        print('Adding Clinical Goals')
        self.label15.Text = "Adding Clinical Goals..."
        ClinicalGoals(self.plan, self.ct, self.sphere_size, self.presc_ld)
        self.update_isodoses()

        self.Close()

    def add_new_plan(self, case, refCT, nob, presc_hd, nof, linac, plan_name):
        # Create a new plan

        plan = case.AddNewPlan(PlanName=plan_name, Comment='SFRT plan',
                               ExaminationName=refCT, AllowDuplicateNames=False)
        orientation = None
        if self.exam.PatientPosition == 'HFP':
            orientation = 'HeadFirstProne'
        if self.exam.PatientPosition == 'HFS':
            orientation = 'HeadFirstSupine'
        if self.exam.PatientPosition == 'FFP':
            orientation = 'FeetFirstProne'
        if self.exam.PatientPosition == 'FFS':
            orientation = 'FeetFirstSupine'
        if orientation is None:
            raise Exception('Patient orientation not recognized')

        # Add Beams
        isocenter_data = {'Position': {'x': self.isocenter.Point.x,
                                       'y': self.isocenter.Point.y,
                                       'z': self.isocenter.Point.z},
                          'NameOfIsocenterToRef': '',
                          'Name': self.isocenter.OfPoi.Name,
                          'Color': self.isocenter.OfPoi.Color, }
        beam_quality = '6'

        beamset_name = self.plan_suffix + "." + self.plan_prefix
        plan.AddNewBeamSet(Name=beamset_name,
                           ExaminationName=refCT,
                           MachineName=linac,
                           Modality='Photons',
                           TreatmentTechnique='VMAT',
                           PatientPosition=orientation,
                           NumberOfFractions=nof,
                           CreateSetupBeams=False,
                           UseLocalizationPointAsSetupIsocenter=False,
                           Comment='')

        beam_set = plan.BeamSets[beamset_name]

        beam_dict = {
            1: {'couch': 0, 'collimator': 85},
            2: {'couch': 0, 'collimator': 125},
            3: {'couch': 0, 'collimator': 50},
            4: {'couch': 10, 'collimator': 57},
            5: {'couch': 350, 'collimator': 312}
        }

        # We start with beam 1, 2 or 3 depending on the number of beams
        start_idx = len(beam_dict) + 1 - nob
        beam_idx = start_idx
        rotation = 'Clockwise'
        past_description = []
        actual_beam_idx = 1
        while beam_idx < 6:
            if rotation == 'Clockwise':
                start = self.GA_start
                stop = self.GA_stop
                rot = "CW"
            else:
                start = self.GA_stop
                stop = self.GA_start
                rot = "CCW"
            beam_description = self.plan_suffix + "_G" + str(start) + rot
            if beam_dict[beam_idx]['couch'] != 0:
                radcalc_couch_angle = (-beam_dict[beam_idx]['couch']) % 360
                beam_description += "_C" + str(radcalc_couch_angle)
            if beam_description in past_description:
                beam_description += "_" + self.plan_suffix + "2"
                while beam_description in past_description:
                    beam_description = beam_description[:-1] + str(int(beam_description[-1]) + 1)
            past_description.append(beam_description)
            beam_set.CreateArcBeam(
                ArcRotationDirection=rotation,
                GantryAngle=start,
                ArcStopGantryAngle=stop,
                Name=self.plan_suffix + str(actual_beam_idx),
                CouchRotationAngle=beam_dict[beam_idx]['couch'],
                CollimatorAngle=beam_dict[beam_idx]['collimator'],
                IsocenterData=isocenter_data,
                BeamQualityId=beam_quality,
                Description=beam_description,
            )
            if rotation == 'Clockwise':
                rotation = 'CounterClockwise'
            else:
                rotation = 'Clockwise'
            if beam_idx == start_idx:
                isocenter_data['NameOfIsocenterToRef'] = isocenter_data['Name']
            beam_idx += 1
            actual_beam_idx += 1

        poi_list = [
            f.OfPoi.Name for f in self.case.PatientModel.StructureSets[refCT].PoiGeometries]
        # Check that ICRU_X exists and raise exception if it doesn't
        if "ICRU_X" not in poi_list:
            # Message box
            MessageBox.Show("ICRU_X not found, run the sphere creation script first")
            raise Exception(
                "ICRU_X not found, run the sphere creation script first")
        icru_x = self.case.PatientModel.StructureSets[refCT].PoiGeometries["ICRU_X"]

        icru_name = 'ICRU_' + self.plan_suffix
        # Check if this point already exists, if so move it to same coordinate as ICRU_X
        if icru_name in poi_list:
            self.case.PatientModel.StructureSets[refCT].PoiGeometries[icru_name].Point = {
                'x': icru_x.Point.x, 'y': icru_x.Point.y, 'z': icru_x.Point.z}
        else:
            self.case.PatientModel.CreatePoi(Name=icru_name,
                                             Examination=self.exam,
                                             Point={'x': icru_x.Point.x,
                                                    'y': icru_x.Point.y,
                                                    'z': icru_x.Point.z},
                                             Color='Blue',
                                             Type='DoseRegion')

        beam_set.AddPoiPrescriptionDoseReference(PoiName=icru_name,
                                                 DoseValue=presc_hd,
                                                 RelativePrescriptionLevel=1.0)

        beam_set.CreateDoseSpecificationPoint(Name='DSP_ICRU_' + self.plan_suffix,
                                              Coordinates={'x': icru_x.Point.x, 'y': icru_x.Point.y, 'z': icru_x.Point.z})

        for beam in beam_set.Beams:
            beam.SetDoseSpecificationPoint(Name='DSP_ICRU_' + self.plan_suffix)

        # Change plan optimization settings
        plan.PlanOptimizations[0].OptimizationParameters.Algorithm.MaxNumberOfIterations = 80
        plan.PlanOptimizations[0].OptimizationParameters.DoseCalculation.IterationsInPreparationsPhase = 30
        plan.PlanOptimizations[0].OptimizationParameters.DoseCalculation.ComputeIntermediateDose = False

        return plan, beam_set

    def create_opt_structures(self, force_delete=False):

        OPT_VTV_HDplus2mm_str = 'OPT_VTV_HDplus2mm'
        PTV_Control_str = 'PTV_Control'
        x_Ring1_str = 'x_Ring1'
        x_Ring2_str = 'x_Ring2'
        x_Ring3_str = 'x_Ring3'
        Skin_str = "Skin"

        roi_list = [f.Name for f in self.case.PatientModel.RegionsOfInterest]

        PTV_LD_str = 'PTV_LD'
        Eval_PTV_LD_str = 'Eval_PTV_LD'

        if PTV_LD_str not in roi_list:
            raise Exception(PTV_LD_str + " not found!")
        VTV_HD_str = 'VTV_HD'

        opt_struct = [OPT_VTV_HDplus2mm_str, PTV_Control_str,
                      x_Ring1_str, x_Ring2_str, x_Ring3_str]

        if force_delete:
            for roi in opt_struct:
                if roi in roi_list:
                    self.case.PatientModel.StructureSets[self.exam.Name].RoiGeometries[roi].DeleteGeometry()
                    roi_list.remove(roi)

        opt_struct_todo = opt_struct
        opt_struct_todo.extend([Skin_str, Eval_PTV_LD_str])

        for roi in opt_struct_todo:
            if roi in roi_list:
                # check if the roi has a contour
                if self.case.PatientModel.StructureSets[self.ct].RoiGeometries[roi].HasContours():
                    opt_struct_todo.remove(roi)

        # Create a wall around the External ROI by 5 mm inwards
        if Skin_str in opt_struct_todo:
            self._wall_roi(self.external_name, -0.5, Skin_str, "Pink", "Organ")

        # Check if "Eval_PTV_LD" exists
        if Eval_PTV_LD_str in opt_struct_todo:
            # Make sure that "PTV_LD" exists
            with CompositeAction('Create Eval_PTV_LD'):
                retval_0 = self._create_roi(Eval_PTV_LD_str, "Blue", "Ptv")
                # Intersection of PTV_LD and External contracted by 5mm
                retval_0.CreateAlgebraGeometry(Examination=self.exam, Algorithm="Auto", ExpressionA={'Operation': "Union", 'SourceRoiNames': [
                    PTV_LD_str], 'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}},
                    ExpressionB={'Operation': "Union", 'SourceRoiNames': [self.external_name],
                                 'MarginSettings': {'Type': "Contract", 'Superior': 0.5, 'Inferior': 0.5, 'Anterior': 0.5, 'Posterior': 0.5, 'Right': 0.5, 'Left': 0.5}},
                    ResultOperation="Intersection", ResultMarginSettings={'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})

        if PTV_Control_str in opt_struct_todo:
            with CompositeAction('ROI algebra (' + PTV_Control_str + ', Image set:  ' + self.ct + ')'):
                retval_0 = self._create_roi(PTV_Control_str, "Green", "Ptv")

                retval_0.CreateAlgebraGeometry(Examination=self.exam, Algorithm="Auto", ExpressionA={'Operation': "Union", 'SourceRoiNames': [Eval_PTV_LD_str], 'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ExpressionB={'Operation': "Union", 'SourceRoiNames': [
                    VTV_HD_str], 'MarginSettings': {'Type': "Expand", 'Superior': 0.5, 'Inferior': 0.5, 'Anterior': 0.5, 'Posterior': 0.5, 'Right': 0.5, 'Left': 0.5}}, ResultOperation="Subtraction", ResultMarginSettings={'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})

                # CompositeAction ends

        if OPT_VTV_HDplus2mm_str in opt_struct_todo:
            with CompositeAction('ROI algebra (' + OPT_VTV_HDplus2mm_str + ', Image set: ' + self.ct + ')'):
                retval_0 = self._create_roi(OPT_VTV_HDplus2mm_str, "128, 0, 64", "Ptv")

                retval_0.CreateAlgebraGeometry(Examination=self.exam, Algorithm="Auto", ExpressionA={'Operation': "Union", 'SourceRoiNames': [VTV_HD_str], 'MarginSettings': {'Type': "Expand", 'Superior': 0.2, 'Inferior': 0.2, 'Anterior': 0.2, 'Posterior': 0.2, 'Right': 0.2, 'Left': 0.2}}, ExpressionB={
                    'Operation': "Union", 'SourceRoiNames': [], 'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ResultOperation="None", ResultMarginSettings={'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})

                # CompositeAction ends

        if x_Ring1_str in opt_struct_todo:
            with CompositeAction('ROI algebra (' + x_Ring1_str + ', Image set:  ' + self.ct + ')'):
                retval_1 = self._create_roi(x_Ring1_str, "138, 74, 11", "Undefined")

                retval_1.CreateAlgebraGeometry(Examination=self.exam, Algorithm="Auto", ExpressionA={'Operation': "Union", 'SourceRoiNames': [Eval_PTV_LD_str], 'MarginSettings': {'Type': "Expand", 'Superior': 1.5, 'Inferior': 1.5, 'Anterior': 1.5, 'Posterior': 1.5, 'Right': 1.5, 'Left': 1.5}}, ExpressionB={'Operation': "Union", 'SourceRoiNames': [
                    Eval_PTV_LD_str], 'MarginSettings': {'Type': "Expand", 'Superior': 0.1, 'Inferior': 0.1, 'Anterior': 0.1, 'Posterior': 0.1, 'Right': 0.1, 'Left': 0.1}}, ResultOperation="Subtraction", ResultMarginSettings={'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})

                # CompositeAction ends

            with CompositeAction('ROI algebra (' + x_Ring1_str + ', Image set: ' + self.ct + ')'):

                retval_1.CreateAlgebraGeometry(Examination=self.exam, Algorithm="Auto", ExpressionA={'Operation': "Union", 'SourceRoiNames': [x_Ring1_str], 'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ExpressionB={'Operation': "Union", 'SourceRoiNames': [
                    self.external_name], 'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ResultOperation="Intersection", ResultMarginSettings={'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})

                # CompositeAction ends

        if x_Ring2_str in opt_struct_todo:
            with CompositeAction('ROI algebra (' + x_Ring2_str + ', Image set:  ' + self.ct + ')'):
                retval_1 = self._create_roi(x_Ring2_str, "81, 11, 138", "Undefined")

                retval_1.CreateAlgebraGeometry(Examination=self.exam, Algorithm="Auto", ExpressionA={'Operation': "Union", 'SourceRoiNames': [Eval_PTV_LD_str], 'MarginSettings': {'Type': "Expand", 'Superior': 3.0, 'Inferior': 3.0, 'Anterior': 3.0, 'Posterior': 3.0, 'Right': 3.0, 'Left': 3.0}}, ExpressionB={'Operation': "Union", 'SourceRoiNames': [
                    Eval_PTV_LD_str], 'MarginSettings': {'Type': "Expand", 'Superior': 1.5, 'Inferior': 1.5, 'Anterior': 1.5, 'Posterior': 1.5, 'Right': 1.5, 'Left': 1.5}}, ResultOperation="Subtraction", ResultMarginSettings={'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})

                # CompositeAction ends

            with CompositeAction('ROI algebra (' + x_Ring2_str + ', Image set:' + self.ct + ')'):

                retval_1.CreateAlgebraGeometry(Examination=self.exam, Algorithm="Auto", ExpressionA={'Operation': "Union", 'SourceRoiNames': [x_Ring2_str], 'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ExpressionB={'Operation': "Union", 'SourceRoiNames': [
                    self.external_name], 'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ResultOperation="Intersection", ResultMarginSettings={'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})

                # CompositeAction ends

        if x_Ring3_str in opt_struct_todo:
            with CompositeAction('ROI algebra (' + x_Ring3_str + ', Image set:  ' + self.ct + ')'):
                retval_1 = self._create_roi(x_Ring3_str, "119, 122, 13", "Undefined")
                retval_1.CreateAlgebraGeometry(Examination=self.exam, Algorithm="Auto", ExpressionA={'Operation': "Union", 'SourceRoiNames': [Eval_PTV_LD_str], 'MarginSettings': {'Type': "Expand", 'Superior': 6.0, 'Inferior': 6.0, 'Anterior': 6.0, 'Posterior': 6.0, 'Right': 6.0, 'Left': 6.0}}, ExpressionB={'Operation': "Union", 'SourceRoiNames': [
                    Eval_PTV_LD_str], 'MarginSettings': {'Type': "Expand", 'Superior': 3.0, 'Inferior': 3.0, 'Anterior': 3.0, 'Posterior': 3.0, 'Right': 3.0, 'Left': 3.0}}, ResultOperation="Subtraction", ResultMarginSettings={'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})

                # CompositeAction ends

            with CompositeAction('ROI algebra (' + x_Ring3_str + ', Image set:  ' + self.ct + ')'):

                retval_1.CreateAlgebraGeometry(Examination=self.exam, Algorithm="Auto", ExpressionA={'Operation': "Union", 'SourceRoiNames': [x_Ring3_str], 'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ExpressionB={'Operation': "Union", 'SourceRoiNames': [
                    self.external_name], 'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}}, ResultOperation="Intersection", ResultMarginSettings={'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})

                # CompositeAction ends

    def _wall_roi(self, roi_name, margin, new_name, new_color="Yellow", roi_type="Organ"):
        # Create a wall around the ROI
        with CompositeAction("Create wall (" + new_name + ", Image set: " + self.exam.Name + ")"):
            roi = self._create_roi(new_name, new_color, roi_type)
            if margin < 0:
                roi.CreateWallGeometry(
                    Examination=self.exam, SourceRoiName=roi_name, OutwardDistance=0, InwardDistance=-margin)
            else:
                roi.CreateWallGeometry(
                    Examination=self.exam, SourceRoiName=roi_name, OutwardDistance=margin, InwardDistance=0)
        return roi

    def _create_roi(self, roi_name, color, roi_type, tissue_name=None, rbe_cell_type_name=None, roi_material=None):
        # Create a new ROI
        roi_list = [f.Name for f in self.case.PatientModel.RegionsOfInterest]
        if roi_name not in roi_list:
            roi = self.case.PatientModel.CreateRoi(
                Name=roi_name, Color=color, Type=roi_type, TissueName=tissue_name, RbeCellTypeName=rbe_cell_type_name, RoiMaterial=roi_material)
        else:
            roi = self.case.PatientModel.RegionsOfInterest[roi_name]
        return roi

    def update_isodoses(self):
        import System.Drawing
        if self.presc_ld == 2000:
            dose_colour_table = {110.0: System.Drawing.Color.FromArgb(250, 102, 0, 102),
                                 107.0: System.Drawing.Color.FromArgb(250, 255, 255, 0),
                                 100.0: System.Drawing.Color.FromArgb(250, 255, 0, 102),
                                 90.0: System.Drawing.Color.FromArgb(250, 0, 0, 128),
                                 85.0: System.Drawing.Color.FromArgb(250, 0, 153, 153),
                                 52.4737: System.Drawing.Color.FromArgb(250, 255, 153, 51),
                                 44.9775: System.Drawing.Color.FromArgb(142, 22, 130),
                                 35.982: System.Drawing.Color.FromArgb(250, 66, 245, 209),
                                 29.985: System.Drawing.Color.FromArgb(250, 102, 102, 255),
                                 28.4857: System.Drawing.Color.FromArgb(250, 0, 153, 51),
                                 14.9925: System.Drawing.Color.FromArgb(250, 128, 64, 0),
                                 7.49625: System.Drawing.Color.FromArgb(250, 255, 255, 255),
                                 0.0: System.Drawing.Color.FromArgb(0, 0, 0, 0)
                                 }
        elif self.presc_ld == 1000:
            dose_colour_table = {110.0: System.Drawing.Color.FromArgb(250, 102, 0, 102),
                                 107.0: System.Drawing.Color.FromArgb(250, 255, 255, 0),
                                 100.0: System.Drawing.Color.FromArgb(250, 255, 0, 102),
                                 90.0: System.Drawing.Color.FromArgb(250, 0, 0, 128),
                                 85.0: System.Drawing.Color.FromArgb(250, 0, 153, 153),
                                 52.4737: System.Drawing.Color.FromArgb(250, 255, 153, 51),
                                 44.9775: System.Drawing.Color.FromArgb(142, 22, 130),
                                 17.9910: System.Drawing.Color.FromArgb(250, 66, 245, 209),
                                 14.9925: System.Drawing.Color.FromArgb(250, 102, 102, 255),
                                 14.2429: System.Drawing.Color.FromArgb(250, 0, 153, 51),
                                 7.49625: System.Drawing.Color.FromArgb(250, 255, 255, 255),
                                 0.0: System.Drawing.Color.FromArgb(0, 0, 0, 0)
                                 }
        # Replace the colour table entirely
        self.case.CaseSettings.DoseColorMap.ColorTable = dose_colour_table
        self.case.CaseSettings.DoseColorMap.ColorMapReferenceType = 'ReferenceValue'
        self.case.CaseSettings.DoseColorMap.PresentationType = 'Absolute'
        self.case.CaseSettings.DoseColorMap.ReferenceValue = self.presc_hd


if __name__ == '__main__':
    form = SFRTForm()
    Application.Run(form)
