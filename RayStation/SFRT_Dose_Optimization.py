from connect import *

import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

from System.Drawing import Point, Color
from System.Windows.Forms import Application, Button, Form, Label, ComboBox, ComboBoxStyle
from System.ComponentModel import BackgroundWorker
from System.Windows import MessageBox

import csv
import os


def find_default_ptv_name(total_roi_list):
    ptv_name = "Eval_PTV_LD"
    if ptv_name not in total_roi_list:
        # Find alternative names
        alternative = ["Eval_PTV_LD", "PTV_LD",
                       "VTV_HD", "VTV_LD", "PTV_Control"]
        for alt in alternative:
            if alt in total_roi_list:
                ptv_name = alt
                break
        if ptv_name not in total_roi_list:
            # Final attempt:
            for roi in total_roi_list:
                if "PTV" in roi:
                    ptv_name = roi
                    break

        if ptv_name not in total_roi_list:
            raise Exception("PTV not found")
    return ptv_name


class Objective(Form):
    def __init__(self):
        self.case = get_current("Case")
        self.exam = get_current("Examination")
        self.plan = get_current("Plan")
        self.beam_set = get_current("BeamSet")
        self.structure_set = self.case.PatientModel.StructureSets[self.exam.Name]

        self.added_obj = 0

        self.total_roi_list = []
        for roi in self.case.PatientModel.RegionsOfInterest:
            self.total_roi_list.append(roi.Name)

        self.Text = 'SFRT Planning Interface'
        self.Height = 300
        self.Width = 320
        self.ForeColor = Color.Black
        self.BackColor = Color.LightGray

        self.sphere_diameter = 1.5
        self.presc_hd = 6670
        self.presc_ld = 2000

        # Label: Prescription dose
        self.label2 = Label()
        self.label2.Text = "Prescription dose"
        self.label2.Location = Point(10, 20)
        self.label2.Height = 20
        self.label2.Width = 170

        # Combobox: Prescription dose
        self.cb_presc = ComboBox()
        self.cb_presc.Parent = self
        self.cb_presc.Location = Point(10, 40)
        self.cb_presc.Items.AddRange(
            ("HD 66.7 Gy / LD 20 Gy", "HD 66.7 Gy / LD 10 Gy"))
        self.cb_presc.SelectedIndex = 0
        self.cb_presc.SelectionChangeCommitted += self.PrescriptionOnChanged
        self.cb_presc.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_presc.Width = 170
        self.PrescriptionOnChanged(None, None)

        self.label3 = Label()
        self.label3.Text = "Sphere diameter"
        self.label3.Location = Point(190, 20)
        self.label3.Height = 20
        self.label3.Width = 100

        # Combobox: Sphere size
        self.cb_sphere = ComboBox()
        self.cb_sphere.Parent = self
        self.cb_sphere.Location = Point(190, 40)
        self.cb_sphere.Items.AddRange(
            ("1.0 cm", "1.5 cm"))
        self.cb_sphere.SelectedIndex = 1
        self.cb_sphere.SelectionChangeCommitted += self.SphereOnChanged
        self.cb_sphere.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_sphere.Width = 80
        self.SphereOnChanged(None, None)

        # Button: Add objectives
        self.button1 = Button()
        self.button1.Text = "Add Step 1"
        self.button1.Location = Point(10, 100)
        self.button1.Click += self.step1_add

        self.button2 = Button()
        self.button2.Text = "Add Step 2"
        self.button2.Location = Point(100, 100)
        self.button2.Click += self.step2_add

        self.button3 = Button()
        self.button3.Text = "Add Step 3"
        self.button3.Location = Point(190, 100)
        self.button3.Click += self.step3_add

        # Button: Run optimization
        self.button7 = Button()
        self.button7.Text = "Run Step 1"
        self.button7.Location = Point(10, 150)
        self.button7.Click += self.step1_run

        self.button8 = Button()
        self.button8.Text = "Run Step 2"
        self.button8.Location = Point(100, 150)
        self.button8.Click += self.step2_run

        self.button9 = Button()
        self.button9.Text = "Run Step 3"
        self.button9.Location = Point(190, 150)
        self.button9.Click += self.step3_run

        self.button4 = Button()
        self.button4.Text = "Final dose"
        self.button4.Location = Point(10, 200)
        self.button4.Click += self.final_click

        self.button5 = Button()
        self.button5.Text = "Pause script"
        self.button5.Location = Point(190, 200)
        self.button5.Click += self.pause_click

        # Status text
        self.label = Label()
        self.label.Text = "Select step number"
        self.label.ForeColor = Color.Black
        self.label.Location = Point(10, 70)
        self.label.Height = 20
        self.label.Width = 400

        self.disable_buttons()
        self.enable_buttons()

        self.Controls.Add(self.label2)
        self.Controls.Add(self.cb_presc)
        self.Controls.Add(self.label3)
        self.Controls.Add(self.cb_sphere)
        self.Controls.Add(self.button1)
        self.Controls.Add(self.button2)
        self.Controls.Add(self.button3)
        self.Controls.Add(self.button4)
        self.Controls.Add(self.button5)
        self.Controls.Add(self.button7)
        self.Controls.Add(self.button8)
        self.Controls.Add(self.button9)
        self.Controls.Add(self.label)
        self.CenterToScreen()

        self.precheck()

    def precheck(self):
        # Check that all optimisation structures are present
        opt_struct = ["VTV_HD", "VTV_LD", "PTV_Control", "GTV", "Skin", "OPT_VTV_HDplus2mm", "Eval_PTV_LD"]
        for roi in opt_struct:
            if roi not in self.total_roi_list:
                MessageBox.Show("ROI " + roi + " not found")
                continue
            if not self.structure_set.RoiGeometries[roi].HasContours():
                MessageBox.Show("ROI " + roi + " has no contours")

    def disable_buttons(self):
        self.button1.Enabled = False
        self.button2.Enabled = False
        self.button3.Enabled = False
        self.button4.Enabled = False
        self.button5.Enabled = False
        self.button7.Enabled = False
        self.button8.Enabled = False
        self.button9.Enabled = False

    def enable_buttons(self):
        self.button1.Enabled = True
        self.button2.Enabled = True
        self.button3.Enabled = True
        self.button4.Enabled = True
        self.button5.Enabled = True

    def SphereOnChanged(self, sender, event):
        if self.cb_sphere.SelectedItem == "1.0 cm":
            self.sphere_diameter = 1.0
        elif self.cb_sphere.SelectedItem == "1.5 cm":
            self.sphere_diameter = 1.5

        self.update_filename()

    def PrescriptionOnChanged(self, sender, event):
        if self.cb_presc.SelectedItem == "HD 66.7 Gy / LD 20 Gy":
            self.presc_hd = 6670
            self.presc_ld = 2000
        elif self.cb_presc.SelectedItem == "HD 66.7 Gy / LD 10 Gy":
            self.presc_hd = 6670
            self.presc_ld = 1000

        self.update_filename()

    def update_filename(self):

        self.filename_dict = {
            1: "SFRT_ObjectiveStep1",
            2: "SFRT_ObjectiveStep2",
            3: "SFRT_ObjectiveStep3",
            "oar": "SFRT_Objective_OAR.csv",
        }
        if self.sphere_diameter == 1.0:
            for i in range(1, 4):
                self.filename_dict[i] = self.filename_dict[i] + "_10mm"
        else:
            for i in range(1, 4):
                self.filename_dict[i] = self.filename_dict[i] + "_15mm"

        if self.presc_ld == 1000:
            for i in range(1, 4):
                self.filename_dict[i] = self.filename_dict[i] + "_10Gy.csv"
        else:
            for i in range(1, 4):
                self.filename_dict[i] = self.filename_dict[i] + "_20Gy.csv"

        # If cannot find the file[1], look in the script's directory
        if not os.path.isfile(self.filename_dict[1]):
            script_dir = os.path.dirname(os.path.abspath(__file__))
            for i in range(1, 4):
                self.filename_dict[i] = os.path.join(script_dir, self.filename_dict[i])
            self.filename_dict["oar"] = os.path.join(script_dir, self.filename_dict["oar"])

    def highlight_button(self, button_obj):
        button_obj.BackColor = Color.LightBlue

    def worker_completed(self, sender, event):
        self.worker.Dispose()
        if event.Error is not None:
            MessageBox.Show(str(event.Error))
        else:
            pass

    def step1_add(self, sender, event):
        self.precheck()
        self._start_worker(self.step1_worker_add)

    def step2_add(self, sender, event):
        self._start_worker(self.step2_worker_add)

    def step3_add(self, sender, event):
        self._start_worker(self.step3_worker_add)

    def _start_worker(self, worker_method):
        self.disable_buttons()
        self.label.ForeColor = Color.Red
        self.worker = BackgroundWorker()
        self.worker.DoWork += worker_method
        self.worker.RunWorkerCompleted += self.worker_completed
        self.worker.RunWorkerAsync()

    def step1_run(self, sender, event):
        self.button7.BackColor = Color.Transparent
        self._start_worker(self.step1_worker_run)

    def step2_run(self, sender, event):
        self.button8.BackColor = Color.Transparent
        self._start_worker(self.step2_worker_run)

    def step3_run(self, sender, event):
        self.button9.BackColor = Color.Transparent
        self._start_worker(self.step3_worker_run)

    def final_click(self, sender, event):
        self._start_worker(self.final_worker)

    def step1_worker_add(self, sender, event):
        self.added_obj = self.add_objectives(
            self.filename_dict[1], self.plan, self.beam_set)
        self.button7.Enabled = True
        self.button5.Enabled = True
        self.label.Text = "Ready to run Step 1 optimization"
        self.label.ForeColor = Color.Green
        self.button7.BackColor = Color.LightBlue

    def step1_worker_run(self, sender, event):
        self.setCustomDoseGrid()

        self.plan.PlanOptimizations[0].RunOptimization()
        self.restrict_MU()
        self.plan.PlanOptimizations[0].ResetOptimization()
        self.plan.PlanOptimizations[0].RunOptimization()

        if not self.check_goals(1, self.added_obj):
            self.plan.PlanOptimizations[0].RunOptimization()

        self.enable_buttons()
        self.Hide()
        await_user_input('Step 1 done, check whether weights need to be adjusted\n'
                         'Press Resume script when ready to add Step 2')
        self.Show()
        self.Focus()

    def step2_worker_add(self, sender, event):
        self.added_obj = self.add_objectives(
            self.filename_dict[2], self.plan, self.beam_set, step2=True, oar_filename=self.filename_dict["oar"])
        self.button8.Enabled = True
        self.button5.Enabled = True
        self.label.Text = "Ready to run Step 2 optimization"
        self.label.ForeColor = Color.Green
        self.button8.BackColor = Color.LightBlue

    def step2_worker_run(self, sender, event):
        self.setCustomDoseGrid()

        self.plan.PlanOptimizations[0].RunOptimization()
        if not self.check_goals(2, self.added_obj):
            self.plan.PlanOptimizations[0].RunOptimization()

        self.enable_buttons()
        self.Hide()
        await_user_input('Step 2 done, check whether weights need to be adjusted\n'
                         'Press Resume script when ready to add Step 3')
        self.Show()
        self.Focus()

    def step3_worker_add(self, sender, event):
        self.added_obj = self.add_objectives(
            self.filename_dict[3], self.plan, self.beam_set)
        self.button9.Enabled = True
        self.button5.Enabled = True
        self.label.Text = "Ready to run Step 3 optimization"
        self.label.ForeColor = Color.Green
        self.button9.BackColor = Color.LightBlue

    def step3_worker_run(self, sender, event):
        self.setCustomDoseGrid()

        self.plan.PlanOptimizations[0].RunOptimization()
        if not self.check_goals(3, self.added_obj):
            self.plan.PlanOptimizations[0].RunOptimization()

        self.enable_buttons()
        self.Hide()
        await_user_input('Step 3 done, check whether weights need to be adjusted\n'
                         'Press Resume script to perform final dose calculation')
        self.Show()
        self.Focus()

    def final_worker(self, sender, event):
        self.setDefaultDoseGrid()

        try:
            with CompositeAction('Final dose calculation'):
                self.beam_set.ComputeDose(
                    ComputeBeamDoses=True, DoseAlgorithm="CCDose", ForceRecompute=False)
        except:
            pass
        self.renormalize_dose()
        self.enable_buttons()
        MessageBox.Show(
            "Final dose calculation complete, you can close the script")

    def pause_click(self, sender, event):
        self.Hide()
        await_user_input('Script paused\n'
                         'Press Resume script to continue')
        self.Show()
        self.Focus()

    def open_csv(self, filename):
        objectives = []
        # CSV is formatted as:
        # ROI, Function, Weight, DoseLevel, PercentVolume
        with open(filename, 'r') as f:
            reader = csv.reader(f)
            # Skip the header
            next(reader)
            for row in reader:
                obj_dict = {}
                obj_dict['ROI'] = row[0]
                obj_dict['Function'] = row[1]
                obj_dict['Weight'] = row[2]
                obj_dict['DoseLevel'] = row[3]
                obj_dict['PercentVolume'] = row[4]
                obj_dict['num_goals'] = int(row[5])
                obj_dict['intermediate_goal'] = []
                for i in range(obj_dict['num_goals']):
                    goal_dict = {}
                    goal_dict['ROI'] = row[6 + 6 * i]
                    goal_dict['Criteria'] = row[7 + 6 * i]
                    goal_dict['Type'] = row[8 + 6 * i]
                    goal_dict['Acceptance'] = row[9 + 6 * i]
                    goal_dict['Parameter'] = row[10 + 6 * i]
                    goal_dict['NewWeight'] = row[11 + 6 * i]
                    obj_dict['intermediate_goal'].append(goal_dict)
                objectives.append(obj_dict)
        return objectives

    def open_intermediate_csv(self, filename):
        goals = []
        with open(filename, 'r') as f:
            reader = csv.reader(f)
            next(reader)
            for row in reader:
                goal_dict = {}
                goal_dict['ROI'] = row[0]
                goal_dict['Criteria'] = row[1]
                goal_dict['Type'] = row[2]
                goal_dict['Acceptance'] = row[3]
                goal_dict['Parameter'] = row[4]
                goal_dict['NewWeight'] = row[5]
                # goal_dict['ObjectiveIdx'] = int(row[6])
                goals.append(goal_dict)
        return goals

    def add_objectives(self, filename, plan, beam_set, step2=False, oar_filename=None):
        beam_set_name = beam_set.DicomPlanLabel
        beam_set_number = plan.BeamSets[beam_set_name].Number - 1
        plan_opt = plan.PlanOptimizations[beam_set_number]

        if hasattr(plan.PlanOptimizations[beam_set_number].Objective, 'ConstituentFunctions'):
            objectives = [
                f for f in plan.PlanOptimizations[beam_set_number].Objective.ConstituentFunctions]
            obj_num = len(objectives)
        else:
            obj_num = 0

        self.start_obj = obj_num
        added_obj = 0
        with CompositeAction('Add Objectives'):
            objectives = self.open_csv(filename)
            if step2:
                oar_list = self.open_csv(oar_filename)
                objectives += oar_list
            for obj in objectives:
                roi_names = obj["ROI"].split(';')
                for roi_name in roi_names:
                    if roi_name in self.total_roi_list:
                        if not self.case.PatientModel.StructureSets[self.exam.Name].RoiGeometries[roi_name].HasContours():
                            continue
                        plan_opt.AddOptimizationFunction(FunctionType=obj["Function"],
                                                         RoiName=roi_name,
                                                         IsConstraint=False,
                                                         RestrictAllBeamsIndividually=False,
                                                         RestrictToBeams=None,
                                                         IsRobust=False,
                                                         RestrictToBeamSet=None)
                        # Add in the weighting factor
                        dose_param = plan_opt.Objective.ConstituentFunctions[
                            obj_num].DoseFunctionParameters
                        dose_param.Weight = float(obj["Weight"])
                        if obj["Function"] == 'MaxDose' or obj["Function"] == 'MinDose':
                            dose_param.DoseLevel = float(obj["DoseLevel"])
                        elif obj["Function"] == 'UniformDose':
                            dose_param.DoseLevel = float(obj["DoseLevel"])
                            dose_param.PercentVolume = 0
                        elif obj["Function"] == 'MaxDVH' or obj["Function"] == 'MinDVH':
                            dose_param.DoseLevel = float(obj["DoseLevel"])
                            if obj["PercentVolume"].endswith("cc"):
                                # The volume is in cc, convert to percentage
                                roi = self.structure_set.RoiGeometries[roi_name]
                                volume = roi.GetRoiVolume()
                                dose_param.PercentVolume = round(float(
                                    obj["PercentVolume"][:-2]) / volume * 100, 3)
                            else:
                                dose_param.PercentVolume = float(
                                    obj["PercentVolume"])
                        elif obj["Function"] == 'MaxEud':
                            dose_param.DoseLevel = float(obj["DoseLevel"])
                            dose_param.EudParameterA = 1.0

                        obj_num += 1
                        added_obj += 1
                        break

        if step2:
            # Only for step2, for any OAR that has been added, if it doesn't have a mean dose objective, add one
            list_of_obj = [
                f.ForRegionOfInterest.Name for f in plan_opt.Objective.ConstituentFunctions]
            for roi in list_of_obj[-added_obj:]:
                list_of_obj_roi_type = [
                    f.DoseFunctionParameters.FunctionType for f in plan_opt.Objective.ConstituentFunctions if f.ForRegionOfInterest.Name == roi]
                if roi not in self.total_roi_list:
                    continue
                # Check that roi is an "Organ" type
                if "Organ" != self.structure_set.RoiGeometries[roi].OfRoi.Type:
                    continue
                if "MaxEud" not in list_of_obj_roi_type:
                    # Evaluate mean dose of the OAR
                    mean_dose = self.plan.TreatmentCourse.TotalDose.GetDoseStatistic(
                        RoiName=roi, DoseType='Average')
                    if mean_dose < 10:
                        continue

                    plan_opt.AddOptimizationFunction(FunctionType='MaxEud',
                                                     RoiName=roi,
                                                     IsConstraint=False,
                                                     RestrictAllBeamsIndividually=False,
                                                     RestrictToBeams=None,
                                                     IsRobust=False,
                                                     RestrictToBeamSet=None)
                    dose_param = plan_opt.Objective.ConstituentFunctions[obj_num].DoseFunctionParameters
                    dose_param.DoseLevel = mean_dose * 0.8  # 80% of last step's mean dose
                    dose_param.EudParameterA = 1.0
                    dose_param.Weight = 5.0

                    added_obj += 1
                    obj_num += 1

        return added_obj

    def find_boundingbox(self):
        roi_list = []
        # Loop through list of ROI in current constraints
        if hasattr(self.plan.PlanOptimizations[0].Objective, 'ConstituentFunctions'):
            for obj in self.plan.PlanOptimizations[0].Objective.ConstituentFunctions:
                if obj.ForRegionOfInterest.Name not in roi_list:
                    roi_list.append(obj.ForRegionOfInterest.Name)

        # Find a way to do the same for the clinical goals
        if hasattr(self.plan.TreatmentCourse.EvaluationSetup, 'EvaluationFunctions'):
            for eval_func in self.plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions:
                if eval_func.ForRegionOfInterest.Name not in roi_list:
                    roi_list.append(eval_func.ForRegionOfInterest.Name)

        if "Skin" in roi_list:
            roi_list.remove("Skin")
        if "Bone" in roi_list:
            roi_list.remove("Bone")
        # Set initial bounding box to be for PTV_LD +/- 5 cm
        # Find if PTV_LD exists
        ptv_name = find_default_ptv_name(self.total_roi_list)

        ptv_bounding_box = self.structure_set.RoiGeometries[ptv_name].GetBoundingBox(
        )
        bounding_box = ptv_bounding_box
        # We only care about z-axis of the bounding box
        bounding_box[0].z = ptv_bounding_box[0].z - 5
        bounding_box[1].z = ptv_bounding_box[1].z + 5

        # For each ROI, find the bounding box and take the union of all bounding boxes
        for roi in roi_list:
            roi = self.structure_set.RoiGeometries[roi]
            current_box = roi.GetBoundingBox()
            bounding_box[0].z = min(bounding_box[0].z, current_box[0].z)
            bounding_box[1].z = max(bounding_box[1].z, current_box[1].z)
        return [bounding_box[0].z, bounding_box[1].z]

    def setCustomDoseGrid(self):
        self.setDefaultDoseGrid()

        bounding_box = self.find_boundingbox()
        # Find current dose grid settings
        dose_grid = self.beam_set.FractionDose.InDoseGrid
        corner = dose_grid.Corner
        voxel_size = dose_grid.VoxelSize
        num_voxels = dose_grid.NrVoxels

        z_voxels = int(
            (bounding_box[1] - bounding_box[0]) / float(voxel_size.z)) + 1

        with CompositeAction('Edit dose grid settings'):
            self.beam_set.UpdateDoseGrid(Corner={'x': corner.x, 'y': corner.y, 'z': bounding_box[0]}, VoxelSize={
                'x': voxel_size.x, 'y': voxel_size.y, 'z': voxel_size.z}, NumberOfVoxels={'x': num_voxels.x, 'y': num_voxels.y, 'z': z_voxels})
            self.beam_set.FractionDose.UpdateDoseGridStructures()

    def setDefaultDoseGrid(self):

        with CompositeAction('Set default grid'):
            self.beam_set.SetDefaultDoseGrid(
                VoxelSize={'x': 0.25, 'y': 0.25, 'z': 0.25})
            self.beam_set.FractionDose.UpdateDoseGridStructures()

    def restrict_MU(self):
        sum_mu = 0
        num_beam = 0
        for beam in self.beam_set.Beams:
            sum_mu += beam.BeamMU
            num_beam += 1
        avg_mu = sum_mu / num_beam
        # subtract MU by 5% and rounddown to closest 50
        avg_mu = avg_mu - 0.05 * avg_mu
        avg_mu = avg_mu - avg_mu % 50
        avg_mu = max(avg_mu, 800)

        for idx, beam in enumerate(self.beam_set.Beams):
            with CompositeAction('Edit beam MU'):
                self.plan.PlanOptimizations[0].OptimizationParameters.TreatmentSetupSettings[0].BeamSettings[idx].ArcConversionPropertiesPerBeam.EditArcBasedBeamOptimizationSettings(
                    CreateDualArcs=False, FinalGantrySpacing=2, MaxArcDeliveryTime=90, BurstGantrySpacing=None, MaxArcMU=avg_mu)

    def _find_existing_goal_index(self, roi_name, goal):
        """Find the index of an existing clinical goal matching the given parameters."""
        eval_funcs = self.plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions
        for idx, eval_func in enumerate(eval_funcs):
            if (eval_func.ForRegionOfInterest.Name == roi_name and
                    eval_func.PlanningGoal.GoalCriteria == goal["Criteria"] and
                    float(eval_func.PlanningGoal.AcceptanceLevel) == float(goal["Acceptance"]) and
                    eval_func.PlanningGoal.Type == goal["Type"] and
                    float(eval_func.PlanningGoal.ParameterValue) == float(goal["Parameter"])):
                return idx
        raise Exception("Goal not found")

    def _matches_objective(self, objective, objective_dict):
        """Check if a plan objective matches a CSV objective row by ROI, function type, and dose level."""
        roi = objective.ForRegionOfInterest.Name
        for roi_name in objective_dict["ROI"].split(';'):
            if roi == roi_name:
                return (objective.DoseFunctionParameters.FunctionType.lower() == objective_dict["Function"].lower() and
                        objective.DoseFunctionParameters.DoseLevel == float(objective_dict["DoseLevel"]))
        return False

    def check_goals(self, step, added_obj):
        beam_set_name = self.beam_set.DicomPlanLabel
        beam_set_number = self.plan.BeamSets[beam_set_name].Number - 1
        plan_opt = self.plan.PlanOptimizations[beam_set_number]
        eval_setup = self.plan.TreatmentCourse.EvaluationSetup

        goal_met = True
        final_obj = self.start_obj + added_obj + 1

        num_goals = 0
        if hasattr(eval_setup, 'EvaluationFunctions'):
            num_goals = len([f for f in eval_setup.EvaluationFunctions])

        objective_list = self.open_csv(self.filename_dict[step])
        if step == 2:
            objective_list += self.open_csv(self.filename_dict["oar"])

        objective_functions = [f for f in plan_opt.Objective.ConstituentFunctions]
        for objective in objective_functions[self.start_obj:final_obj]:
            for objective_dict in objective_list:
                if not self._matches_objective(objective, objective_dict):
                    continue

                for goal in objective_dict["intermediate_goal"]:
                    # Find the ROI name that exists in the current structure set
                    roi_obj = None
                    for candidate in goal["ROI"].split(';'):
                        if candidate in self.total_roi_list:
                            roi_obj = candidate
                            break

                    # Try to add the clinical goal; if it already exists, find its index
                    try:
                        with CompositeAction('Add clinical goal'):
                            eval_setup.AddClinicalGoal(
                                RoiName=roi_obj,
                                GoalCriteria=goal["Criteria"],
                                AcceptanceLevel=goal["Acceptance"],
                                GoalType=goal["Type"],
                                ParameterValue=goal["Parameter"],
                                IsComparativeGoal=False
                            )
                        goal_idx = num_goals
                        newly_added = True
                    except Exception as e:
                        print(e)
                        goal_idx = self._find_existing_goal_index(roi_obj, goal)
                        newly_added = False

                    if not eval_setup.EvaluationFunctions[goal_idx].EvaluateClinicalGoal():
                        print("Intermediate goal not met")
                        objective.DoseFunctionParameters.Weight = goal["NewWeight"]
                        goal_met = False
                    else:
                        print("Intermediate goal met")

                    # Remove the clinical goal if it was newly added
                    if newly_added:
                        eval_setup.DeleteClinicalGoal(
                            FunctionToRemove=eval_setup.EvaluationFunctions[num_goals])

        return goal_met

    def renormalize_dose(self):
        frame_of_ref = self.beam_set.FrameOfReference
        dsp_point = [x for x in self.beam_set.DoseSpecificationPoints][0]
        dsp_coord = {'x': dsp_point.Coordinates.x,
                     'y': dsp_point.Coordinates.y, 'z': dsp_point.Coordinates.z}
        dose = self.plan.TreatmentCourse.TotalDose.InterpolateDoseInPoint(
            Point=dsp_coord, PointFrameOfReference=frame_of_ref)
        prescription = self.beam_set.Prescription.PrimaryPrescriptionDoseReference.DoseValue
        factor = float(prescription) / dose
        try:
            self.beam_set.Prescription.PrescriptionDoseReferences[0].RelativePrescriptionLevel = round(
                factor, 3)
        except:
            # This might fail if factor is >2 or <0.1 but this is just for the display, we don't actually renormalize.
            pass


if __name__ == '__main__':
    objective_form = Objective()
    Application.Run(objective_form)
