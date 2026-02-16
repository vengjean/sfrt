# -*- coding: utf-8 -*-

# Created in April 2024

# @author: Veng Jean Heng

import clr
import sys

from connect import *

import platform

if platform.python_implementation() != "CPython":
    print("Python interpreter should be CPython, but is currently %s" %
          (platform.python_implementation()))
    sys.exit()

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
import System
from System import Windows
from System.Windows import Window, Application, MessageBox
from System.Windows.Forms import StatusBar
from System.Windows.Markup import XamlReader
from System.IO import StringReader
from System.Xml import XmlReader
from System.Threading import Thread, ThreadStart, ApartmentState

# background worker
from System.ComponentModel import BackgroundWorker

from System.Collections.ObjectModel import ObservableCollection
from System.Windows import LogicalTreeHelper as lth

from connect import *

# xaml for GUI
xaml = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SFRT Contour creation tool"
        Margin="10,10,10,10"
        SizeToContent="WidthAndHeight"
        WindowStartupLocation="Manual">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="*" />
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="*" />
            <RowDefinition Height="*" />
            <RowDefinition Height="*" />
            <RowDefinition Height="10" />
            <RowDefinition Height="*" />
            <RowDefinition Height="*" />
            <RowDefinition Height="*" />
            <RowDefinition Height="*" />
            <RowDefinition Height="*" />
            <RowDefinition Height="*" />
            <RowDefinition Height="30" />
            <RowDefinition Height="*" />
            <RowDefinition Height="*" />
            <RowDefinition Height="*" />
            <RowDefinition Height="10" />
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0"
                   Grid.Column="0"
                   Margin="5"
                   Text="Select examination" />
        <ComboBox Name="SelectExaminationComboBox"
                  Grid.Row="0"
                  Grid.Column="1"
                  Grid.ColumnSpan="2"
                  Margin="5" />
        <TextBlock Grid.Row="1"
                   Grid.Column="0"
                   Margin="5"
                   Text="Select GTVm" />
        <ComboBox Name="SelectGTVComboBox"
                  Grid.Row="1"
                  Grid.Column="1"
                  Grid.ColumnSpan="2"
                  Margin="5" />
        <TextBlock Grid.Row="2"
                   Grid.Column="0"
                   Margin="5"
                   Text="Select PTVm_2000" />
        <ComboBox Name="SelectPTVComboBox"
                  Grid.Row="2"
                  Grid.Column="1"
                  Grid.ColumnSpan="2"
                  Margin="5" />
        <TextBlock Grid.Row="3"
                   Grid.Column="0"
                   Margin="5"
                   Text="Select External contour" />
        <ComboBox Name="SelectExternalComboBox"
                  Grid.Row="3"
                  Grid.Column="1"
                  Grid.ColumnSpan="2"
                  Margin="5" />
        <TextBlock Grid.Row="5"
                   Grid.Column="0"
                   Margin="5"
                   Text="Select ROI" />
        <TextBlock Grid.Row="5"
                   Grid.Column="1"
                   Margin="5"
                   Text="Enter margin [cm]" />
        <ComboBox Name="SelectRoiComboBox"
                  Grid.Row="6"
                  Grid.Column="0"
                  Width="150"
                  Margin="5"
                  VerticalAlignment="Center" />
        <TextBox Name="margin"
                 Grid.Row="6"
                 Grid.Column="1"
                 Text="2.0"
                 Margin="5" />
        <Button Grid.Row="6"
                Grid.Column="3"
                Width="100"
                Margin="5"
                Name="button_click_add"
                Content="Add" />
        <Button Grid.Row="7"
                Grid.Column="3"
                Width="100"
                Margin="5"
                Name="button_click_remove"
                Content="Remove" />
        <DataGrid  Grid.Row="8" Grid.Column="0"  Grid.ColumnSpan="2" Grid.RowSpan="4" HorizontalAlignment="Stretch"
                Margin="5" SelectionMode="Extended" SelectionUnit= "FullRow" Height="200"
                Name="roiList" AutoGenerateColumns="False" CanUserSortColumns="True">
            <DataGrid.Columns>
            <DataGridTemplateColumn Header="Name" IsReadOnly="True" Width="200" SortMemberPath="name_clr">
                <DataGridTemplateColumn.CellTemplate>
                <DataTemplate>
                    <TextBlock Margin="2" Text="{Binding Path=name_clr}">
                    </TextBlock>
                </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
            <DataGridTemplateColumn Header="margin [cm]" IsReadOnly="True" Width="100" SortMemberPath="margin_clr">
                <DataGridTemplateColumn.CellTemplate>
                <DataTemplate>
                    <TextBlock Margin="2" Text="{Binding Path=margin_clr}">
                    </TextBlock>
                </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
            </DataGrid.Columns>
            </DataGrid>
        <TextBlock Grid.Row="12"
                   Grid.Column="0"
                   Margin="5"
                   Text="Select lattice configuration" />
        <ComboBox Name="SelectLatticeComboBox"
                  Grid.Row="12"
                  Grid.Column="1"
                  Grid.ColumnSpan="2"
                  Margin="5" />
        <Button Grid.Row="13"
                Grid.Column="0"
                Margin="5"
                Name="button_click_generate"
                Content="Generate contours" />
        <TextBlock Grid.Row="13"
                   Grid.Column="1"
                   Margin="5"
                   Text="This process will take several minutes." />
        <StatusBar Grid.Row="14"
                     Grid.Column="0"
                     Grid.ColumnSpan="4"
                     HorizontalAlignment="Stretch">
                <StatusBarItem>
                 <TextBlock Name="statusText" 
                 HorizontalAlignment="Center"
                 />
                </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
"""
xr = XmlReader.Create(StringReader(xaml))

error_message_header = "Missing information / setup"

dose_algorithm = "Undefined"


class RoiGeometry(System.Object):
    """Presentation object for ROIs"""
    __namespace__ = "DotNet"

    def init(self, rg_name, margin):
        self.rg_name = rg_name
        self.margin = margin

    @clr.clrproperty(str)
    def name_clr(self):
        return self.rg_name

    @clr.clrproperty(str)
    def margin_clr(self):
        return str(self.margin)


class MyWindow(Window):
    def __init__(self):
        self.window = XamlReader.Load(xr)
        self.patient = get_current('Patient')
        self.case = get_current('Case')
        self.examination_names = System.Collections.Generic.List[str]()
        lattice_config_names = System.Collections.Generic.List[str]()
        lattice_config = ["Default",
                          "1.5 cm diameter spheres", "1.0 cm diameter spheres"]
        for config_name in lattice_config:
            lattice_config_names.Add(config_name)
        for exam_name in [exam.Name for exam in self.case.Examinations]:
            self.examination_names.Add(exam_name)

        lth.FindLogicalNode(
            self.window, "SelectExaminationComboBox").ItemsSource = self.examination_names
        lth.FindLogicalNode(
            self.window, "SelectExaminationComboBox").SelectedIndex = 0
        lth.FindLogicalNode(
            self.window, "SelectExaminationComboBox").SelectionChanged += self.SelectedExaminationChanged

        lth.FindLogicalNode(self.window, "SelectRoiComboBox").SelectedIndex = 0

        lth.FindLogicalNode(
            self.window, "SelectLatticeComboBox").ItemsSource = lattice_config_names
        lth.FindLogicalNode(
            self.window, "SelectLatticeComboBox").SelectedIndex = 0
        lth.FindLogicalNode(
            self.window, "SelectLatticeComboBox").SelectionChanged += self.SelectedLatticeChanged
        self.radius = 0.75
        self.xy_spacing = 6.0
        self.z_spacing = 3.0
        self.threshold = 6

        lth.FindLogicalNode(
            self.window, "button_click_add").Click += self.button_click_add
        lth.FindLogicalNode(
            self.window, "button_click_generate").Click += self.button_click_generate
        lth.FindLogicalNode(
            self.window, "button_click_remove").Click += self.button_click_remove

        # Holds the ROI(s) and margin value(s) selected
        self.roi_dictionary = {}

        self.update_list_of_geometries()

        Application().Run(self.window)

    def check_ready(self):
        # Check if all the necessary selections have been made
        # Update the status text to ready if so
        if lth.FindLogicalNode(self.window, "SelectExaminationComboBox").SelectedItem is None:
            self.update_status("Select an examination")
            return False
        if lth.FindLogicalNode(self.window, "SelectGTVComboBox").SelectedItem is None:
            self.update_status("Select a GTV")
            return False
        if lth.FindLogicalNode(self.window, "SelectPTVComboBox").SelectedItem is None:
            self.update_status("Select a PTV")
            return False
        if lth.FindLogicalNode(self.window, "SelectExternalComboBox").SelectedItem is None:
            self.update_status("Select an external contour")
            return False
        if lth.FindLogicalNode(self.window, "SelectLatticeComboBox").SelectedItem is None:
            self.update_status("Select a lattice configuration")
            return False
        self.update_status("Ready")

    # Select examination
    def SelectedExaminationChanged(self, sender, event):
        examination_name = sender.SelectedItem
        self.update_list_of_geometries()

        self.roi_dictionary = {}

    def SelectedGTVChanged(self, sender, event):
        self.gtv_name = sender.SelectedItem
        self.check_ready()

    def SelectedPTVChanged(self, sender, event):
        self.ptv_name = sender.SelectedItem
        self.check_ready()

    def SelectedExternalChanged(self, sender, event):
        self.external_name = sender.SelectedItem
        self.check_ready()

    def SelectedLatticeChanged(self, sender, event):
        lattice_name = sender.SelectedItem
        if lattice_name == "Default":
            self.radius = 0.75
            self.xy_spacing = 6.0
            self.z_spacing = 3.0
            self.threshold = 6
        elif lattice_name == "1.5 cm diameter spheres":
            self.radius = 0.75
            self.xy_spacing = 6.0
            self.z_spacing = 3.0
            self.threshold = 0
        elif lattice_name == "1.0 cm diameter spheres":
            self.radius = 0.5
            self.xy_spacing = 4.0
            self.z_spacing = 2.0
            self.threshold = 0

    def update_list_oar(self):
        oar_names = System.Collections.Generic.List[str]()
        for r in self.structure_set.RoiGeometries:
            if r.PrimaryShape is not None:
                if r.OfRoi.Type == "Organ" and r.OfRoi.Name not in self.roi_dictionary:
                    oar_names.Add(r.OfRoi.Name)

        lth.FindLogicalNode(
            self.window, "SelectRoiComboBox").ItemsSource = oar_names
        lth.FindLogicalNode(self.window, "SelectRoiComboBox").SelectedIndex = 0

    def update_list_of_geometries(self):
        examination_name = lth.FindLogicalNode(
            self.window, "SelectExaminationComboBox").SelectedItem
        self.structure_set = self.case.PatientModel.StructureSets[examination_name]
        roi_names = System.Collections.Generic.List[str]()
        gtv_names = System.Collections.Generic.List[str]()
        ptv_names = System.Collections.Generic.List[str]()
        oar_names = System.Collections.Generic.List[str]()
        for r in self.structure_set.RoiGeometries:
            if r.PrimaryShape is not None:
                roi_names.Add(r.OfRoi.Name)
                if r.OfRoi.Type == "Gtv":
                    gtv_names.Add(r.OfRoi.Name)
                if r.OfRoi.Type == "Ptv":
                    ptv_names.Add(r.OfRoi.Name)
                if r.OfRoi.Type == "Organ" and r.OfRoi.Name not in self.roi_dictionary:
                    oar_names.Add(r.OfRoi.Name)

        lth.FindLogicalNode(
            self.window, "SelectRoiComboBox").ItemsSource = oar_names
        lth.FindLogicalNode(
            self.window, "SelectGTVComboBox").ItemsSource = gtv_names
        lth.FindLogicalNode(
            self.window, "SelectGTVComboBox").SelectionChanged += self.SelectedGTVChanged

        # Set default to first item in list with name "GTVm"
        if "GTVm" in roi_names:
            lth.FindLogicalNode(
                self.window, "SelectGTVComboBox").SelectedItem = "GTVm"
            self.gtv_name = "GTVm"

        lth.FindLogicalNode(
            self.window, "SelectPTVComboBox").ItemsSource = ptv_names
        lth.FindLogicalNode(
            self.window, "SelectPTVComboBox").SelectionChanged += self.SelectedPTVChanged

        # Set default to first item in list with name "PTVm_2000"
        if "Eval_PTVm_2000" in roi_names:
            lth.FindLogicalNode(
                self.window, "SelectPTVComboBox").SelectedItem = "Eval_PTVm_2000"
            self.ptv_name = "Eval_PTVm_2000"
        elif "PTVm_2000" in roi_names:
            lth.FindLogicalNode(
                self.window, "SelectPTVComboBox").SelectedItem = "PTVm_2000"
            self.ptv_name = "PTVm_2000"

        lth.FindLogicalNode(
            self.window, "SelectExternalComboBox").ItemsSource = roi_names
        lth.FindLogicalNode(
            self.window, "SelectExternalComboBox").SelectionChanged += self.SelectedExternalChanged
        # Set default to first item in list with name "External"
        if "External" in roi_names:
            lth.FindLogicalNode(
                self.window, "SelectExternalComboBox").SelectedItem = "External"
            self.external_name = "External"

        self.check_ready()

    # Add ROI with margin to dictionary
    def button_click_add(self, sender, e):
        if not lth.FindLogicalNode(self.window, "SelectRoiComboBox").SelectedItem:
            return
        roi_name = lth.FindLogicalNode(
            self.window, "SelectRoiComboBox").SelectedItem
        margin_value = lth.FindLogicalNode(self.window, "margin").Text
        try:
            if float(margin_value) > 9:
                MessageBox.Show('Warning: Check margin value')
            if roi_name in self.roi_dictionary:
                MessageBox.Show('ROI already selected')
            elif float(margin_value) > 0 and (roi_name != None or roi_name != ''):
                self.roi_dictionary.update({roi_name: float(margin_value)})
                roi_info_str = roi_name + ': ' + margin_value
            else:
                MessageBox.Show('ROI with margin value could not be selected')
        except:
            MessageBox.Show('Margin value must be a number')

        self.update_roi_dictionary_display()
        self.update_list_oar()

    # Remove ROI with margin from dictionary
    def button_click_remove(self, sender, e):
        if not lth.FindLogicalNode(self.window, "roiList").SelectedItem:
            return
        roi_geometry = lth.FindLogicalNode(self.window, "roiList").SelectedItem
        del self.roi_dictionary[roi_geometry.rg_name]
        self.update_roi_dictionary_display()
        self.update_list_oar()

    def update_roi_dictionary_display(self):
        roi_margins = System.Collections.Generic.List[RoiGeometry]()
        for dict_item in self.roi_dictionary:
            rg = RoiGeometry()
            rg.init(dict_item, self.roi_dictionary[dict_item])
            roi_margins.Add(rg)
        lth.FindLogicalNode(self.window, "roiList").ItemsSource = roi_margins

    def button_click_generate(self, sender, e):
        self.update_status(
            'Contour generation in progress...This may take several minutes')
        params = {
            "case": self.case,
            "examination": self.examination_names[lth.FindLogicalNode(self.window, "SelectExaminationComboBox").SelectedIndex],
            "body_name": self.external_name,
            "body_margin": 1.8,
            "oar_dict": self.roi_dictionary,
            "gtv_name": self.gtv_name,
            "ptv_name": self.ptv_name,
            "xy_spacing": self.xy_spacing,
            "z_spacing": self.z_spacing,
            "radius": self.radius,
            "threshold": self.threshold,
        }
        # Spawn backgroundworker
        self.worker = BackgroundWorker()
        generate_obj = SFRT_Contour(params)
        self.worker.DoWork += generate_obj.generate_contours
        self.worker.RunWorkerCompleted += self.worker_completed
        self.worker.RunWorkerAsync()

    def worker_completed(self, sender, e):
        self.worker.Dispose()
        if e.Error:
            MessageBox.Show(str(e.Error))
            self.update_status('Error in contour generation')
        else:
            # MessageBox.Show('Contour generation complete')
            self.update_status('Contour generation complete')
            self.window.Close()

    def update_status(self, message):
        lth.FindLogicalNode(self.window, "statusText").Text = message
        if message == "Ready" or message == "Contour generation complete":
            lth.FindLogicalNode(
                self.window, "statusText").Background = System.Windows.Media.Brushes.LightGreen
            lth.FindLogicalNode(
                self.window, "button_click_generate").IsEnabled = True
        elif "in progress" in message:
            lth.FindLogicalNode(
                self.window, "statusText").Background = System.Windows.Media.Brushes.LightYellow
            lth.FindLogicalNode(
                self.window, "button_click_generate").IsEnabled = False
        else:
            lth.FindLogicalNode(
                self.window, "statusText").Background = System.Windows.Media.Brushes.LightCoral
            lth.FindLogicalNode(
                self.window, "button_click_generate").IsEnabled = False


class Structure():
    def __init__(self, attrs):
        self.case = attrs["case"]
        self.examination = attrs["examination"]
        self.image_set = self.examination.Name
        self.structure_set = self.case.PatientModel.StructureSets[self.image_set]
        struct_name = attrs["name"]
        print(struct_name)
        self.structure = self.structure_set.RoiGeometries[struct_name]
        if not hasattr(self.structure.PrimaryShape, "Contours"):
            self.structure.SetRepresentation(Representation="Contours")
        self.contours = self.structure.PrimaryShape.Contours
        self.initialize_polygon()

    def initialize_polygon(self):
        # Make a list of polygons for each contour
        self.polygons = []
        z_coords = []
        for contour in self.contours:
            polygon = []
            for point in contour:
                polygon.append([point.x, point.y])
            slice_z = round(contour[0].z, 3)
            if slice_z in z_coords:
                idx = z_coords.index(slice_z)
                self.polygons[idx].append(polygon)
            else:
                self.polygons.append([polygon])
                z_coords.append(round(contour[0].z, 3))
        # Sort the z_coords and sort polygons accordingly
        idx = np.argsort(z_coords)
        self.polygons = [self.polygons[i] for i in idx]
        self.z_coords = [z_coords[i] for i in idx]

    def points_in_polygon(self, polygon, pts):
        # Found from stackoverflow: https://stackoverflow.com/a/67460792
        # To circumvent having to use matplotlib.path
        pts = np.asarray(pts, dtype='float32')
        polygon = np.asarray(polygon, dtype='float32')
        contour2 = np.vstack((polygon[1:], polygon[:1]))
        test_diff = contour2 - polygon
        mask1 = (pts[:, None] == polygon).all(-1).any(-1)
        m1 = (polygon[:, 1] > pts[:, None, 1]) != (
            contour2[:, 1] > pts[:, None, 1])
        slope = ((pts[:, None, 0] - polygon[:, 0]) * test_diff[:, 1]) - \
            (test_diff[:, 0] * (pts[:, None, 1] - polygon[:, 1]))
        m2 = slope == 0
        mask2 = (m1 & m2).any(-1)
        m3 = (slope < 0) != (contour2[:, 1] < polygon[:, 1])
        m4 = m1 & m3
        count = np.count_nonzero(m4, axis=-1)
        mask3 = ~(count % 2 == 0)
        mask = mask1 | mask2 | mask3
        return mask

    def get_mask(self, coords):
        # Return a mask of the target on the coords grid
        mask = np.zeros(coords["num_voxels"], dtype=bool)
        # Loop over each slice
        for idx, z in enumerate(coords["z_voxels"]):
            # find the closest slice_z to z within slice_spacing
            slice_idx = np.argmin(np.abs(self.z_coords - z))
            if np.abs(self.z_coords[slice_idx] - z) > coords["spacing"][2]:
                continue
            x_indices, y_indices = np.indices(
                (coords["num_voxels"][0], coords["num_voxels"][1]))
            # Flatten the grid and stack x and y coordinates
            flat = np.stack((coords["grid"][0, y_indices, x_indices].flatten(
            ), coords["grid"][1, y_indices, x_indices].flatten()), axis=-1)
            flat_mask = np.zeros(flat.shape[0], dtype=bool)
            for polygon in self.polygons[slice_idx]:
                section_mask = self.points_in_polygon(polygon, flat)
                flat_mask = np.logical_or(flat_mask, section_mask)
            # Reshape the mask back to the original shape
            mask[:, :, idx] = flat_mask.reshape(
                coords["num_voxels"][0], coords["num_voxels"][1])

        return mask


class SFRT_Contour():
    def __init__(self, attrs):
        self.patient = get_current("Patient")
        self.case = attrs["case"]
        self.exam = self.case.Examinations[attrs["examination"]]
        print(self.exam.Name)
        self.exam.SetPrimary()
        self.series = self.exam.Series[0]
        self.slice_spacing = abs(
            self.series.ImageStack.SlicePositions[1] - self.series.ImageStack.SlicePositions[0])
        self.structure_set = self.case.PatientModel.StructureSets[self.exam.Name]

        self.roi_list = [f.Name for f  in self.case.PatientModel.RegionsOfInterest]

        self.xy_spacing = attrs.get("xy_spacing", 6.0)
        self.z_spacing = attrs.get("z_spacing", 3.0)
        self.radius = attrs.get("radius", 0.75)
        self.params = {
            "body_name": attrs.get("body_name", "External"),
            "body_margin": attrs.get("body_margin", 2.0),
            "oar_dict": attrs.get("oar_dict", {}),
            "gtv_name": attrs.get("gtv_name", "GTVm"),
            "ptv_name": attrs.get("ptv_name", "PTVm"),
            "sphere_radius": self.radius,
        }
        self.threshold = attrs.get("threshold", 0)

    def check_existing_roi(self):
        list_roi_to_check = ["GTVm_core", "x_PTVm_core", "x_OARs_all", "Initial_hot_spheres", "Initial_cold_spheres", "x_external_shrink"]
        list_geometry_to_check = ["x_OARs_PRV", "PTVm_6670_1.5, Eval_PTVm_Avoid_1.5", "PTVm_6670_1.0, Eval_PTVm_Avoid_1.0"]
        # This is not clean. We should just check the list of existing roi and delete if it matches list above
        # But for some reason the list of existing roi is not properly updated?? It keeps missing some roi
        for roi in list_roi_to_check:
            try:
                self._delete_roi(roi)
            except:
                pass
        for roi in list_geometry_to_check:
            try:
                self._delete_geometry(roi)
            except:
                pass

    def generate_contours(self, sender, e):
        self.check_existing_roi()
        gtv_core, ptv_core = self.create_substructures(self.params)

        enough_spheres = self.make_ptv_spheres(gtv_core.Name, ptv_core.Name)
        if not enough_spheres:
            self._delete_geometry("x_OARs_PRV")
            self._delete_roi("x_OARs_all")
            self._delete_roi("GTVm_core")
            self._delete_roi("x_PTVm_core")
            self.radius = 0.5
            self.params["sphere_radius"] = self.radius
            self.xy_spacing = 4.0
            self.z_spacing = 2.0

            self.threshold = 0
            gtv_core, ptv_core = self.create_substructures(self.params)
            enough_spheres = self.make_ptv_spheres(
                gtv_core.Name, ptv_core.Name)
            # if not enough_spheres:
            #     self._delete_roi("x_OARs_PRV")
            #     self._delete_roi("x_OARs_all")
            #     self._delete_roi("GTVm_core")
            #     self._delete_roi("x_PTVm_core")
            #     raise Exception("Not enough spheres")

        self._delete_roi("x_OARs_all")
        self._delete_roi("GTVm_core")
        self._delete_roi("x_PTVm_core")
        return True

    def _create_roi(self, roi_name, color="Yellow", roi_type="Undefined"):
        # Create a new ROI
        roi_list = [f.Name for f  in self.case.PatientModel.RegionsOfInterest]
        if roi_name not in roi_list:
            roi = self.case.PatientModel.CreateRoi(
                Name=roi_name, Color=color, Type=roi_type, TissueName=None, RbeCellTypeName=None, RoiMaterial=None)
        else:
            roi = self.case.PatientModel.RegionsOfInterest[roi_name]
        return roi

    def _wall_roi(self, roi_name, margin, new_name, new_color="Yellow", roi_type="Undefined"):
        # Create a wall around the ROI
        with CompositeAction(f"Create wall ({new_name}, Image set: {self.exam.Name})"):
            roi = self._create_roi(new_name, new_color, roi_type)
            if margin < 0:
                roi.CreateWallGeometry(
                    Examination=self.exam, SourceRoiName=roi_name, OutwardDistance=0, InwardDistance=-margin)
            else:
                roi.CreateWallGeometry(
                    Examination=self.exam, SourceRoiName=roi_name, OutwardDistance=margin, InwardDistance=0)
        return roi

    def _expand_roi(self, roi_name, margin, new_name, color="Yellow", roi_type="Undefined"):
        # Expand ROI by margin
        with CompositeAction(f"Expand ROI ({new_name}, Image set: {self.exam.Name})"):
            if new_name != roi_name:
                roi = self._create_roi(new_name, color, roi_type)
            else:
                roi = self.case.PatientModel.RegionsOfInterest[roi_name]
            roi.CreateMarginGeometry(
                Examination=self.exam,
                SourceRoiName=roi_name,
                MarginSettings={'Type': "Expand", 'Superior': margin, 'Inferior': margin,
                                'Anterior': margin, 'Posterior': margin, 'Right': margin, 'Left': margin}
            )
        return roi

    def _contract_roi(self, roi_name, margin, new_name, color="Yellow", roi_type="Undefined"):
        # Contract ROI by margin
        with CompositeAction(f"Shrink ROI ({new_name}, Image set: {self.exam.Name})"):
            if new_name != roi_name:
                roi = self._create_roi(new_name, color, roi_type)
            else:
                roi = self.case.PatientModel.RegionsOfInterest[new_name]
            roi.CreateMarginGeometry(
                Examination=self.exam,
                SourceRoiName=roi_name,
                MarginSettings={'Type': "Contract", 'Superior': margin, 'Inferior': margin,
                                'Anterior': margin, 'Posterior': margin, 'Right': margin, 'Left': margin}
            )
        return roi

    def _delete_roi(self, roi_name):
        # Delete the ROI
        self.case.PatientModel.RegionsOfInterest[roi_name].DeleteRoi()

    def _delete_geometry(self, roi_name):
        # Delete the geometry of the ROI
        self.case.PatientModel.StructureSets[self.exam.Name].RoiGeometries[roi_name].DeleteGeometry()

    def _simplify_roi(self, roi_name):

        self.structure_set.SimplifyContours(
            RoiNames=[roi_name],
            RemoveHoles3D=False,
            RemoveSmallContours=True,
            AreaThreshold=0.1,
            ReduceMaxNumberOfPointsInContours=False,
            MaxNumberOfPoints=None,
            CreateCopyOfRoi=False,
            ResolveOverlappingContours=False
        )

    def _subtract_roi(self, roi1, roi2, new_name, color="Yellow", roi_type="Undefined"):
        # Subtract ROI2 from ROI1
        with CompositeAction(f"Subtract ROI ({new_name}, Image set: {self.exam.Name})"):
            if new_name != roi1 and new_name != roi2:
                roi = self._create_roi(new_name, color, roi_type)
            else:
                roi = self.case.PatientModel.RegionsOfInterest[new_name]
            roi.CreateAlgebraGeometry(
                Examination=self.exam,
                Algorithm="Auto",
                ExpressionA={
                    'Operation': "Union",
                    'SourceRoiNames': [roi1],
                    'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}
                },
                ExpressionB={'Operation': "Union",
                             'SourceRoiNames': [roi2],
                             'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}
                             },
                ResultOperation="Subtraction",
                ResultMarginSettings={'Type': "Expand", 'Superior': 0, 'Inferior': 0,
                                      'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}
            )
        return roi

    def _intersection_roi(self, roi1, roi2, new_name, color="Yellow", roi_type="Undefined"):
        # Intersect ROI1 with ROI2
        with CompositeAction(f"Intersect ROI ({new_name}, Image set: {self.exam.Name})"):
            if new_name != roi1 and new_name != roi2:
                roi = self._create_roi(new_name, color, roi_type)
            else:
                roi = self.case.PatientModel.RegionsOfInterest[new_name]
            roi.CreateAlgebraGeometry(
                Examination=self.exam,
                Algorithm="Auto",
                ExpressionA={
                    'Operation': "Union",
                    'SourceRoiNames': [roi1],
                    'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}
                },
                ExpressionB={'Operation': "Union",
                             'SourceRoiNames': [roi2],
                             'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}
                             },
                ResultOperation="Intersection",
                ResultMarginSettings={'Type': "Expand", 'Superior': 0, 'Inferior': 0,
                                      'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}
            )
        return roi

    def create_substructures(self, params):

        body_name = params["body_name"]
        body_margin = params["body_margin"]
        oar_dict = params["oar_dict"]
        gtv_name = params["gtv_name"]
        ptv_name = params["ptv_name"]
        sphere_radius = params["sphere_radius"]

        gtv_margin = 0.5

        # Create a wall around the body
        combined_prv = self._wall_roi(body_name, -body_margin, "x_OARs_PRV")

        # Create a combined OAR PRV structure
        oar_combined = self._create_roi(
            "x_OARs_all", "Pink", "Undefined")

        external_shrink = self._contract_roi(
            body_name, 0.5, "x_external_shrink", "Pink", "Undefined")

        for oar, margin in oar_dict.items():
            self._expand_roi(oar, margin, f"x_{oar}_temp")
            self._union_roi(combined_prv.Name, f"x_{oar}_temp")
            self._delete_roi(f"x_{oar}_temp")

            self._union_roi(oar_combined.Name, oar)

        # Shrink the GTVm_core by gtv_margin
        gtv_core = self._contract_roi(
            gtv_name, gtv_margin, "GTVm_core", "Maroon", "Gtv")

        # Perform subtraction of GTV from the combined PRV
        gtv_core = self._subtract_roi(
            "GTVm_core", combined_prv.Name, "GTVm_core", "Maroon", "Gtv")
        # Shrink the GTVm_core by sphere_radius
        gtv_core = self._contract_roi(
            "GTVm_core", sphere_radius, "GTVm_core", "Maroon", "Gtv")
        self._simplify_roi(gtv_core.Name)

        # Perform subtraction of PTV from the combined OAR
        ptv_core = self._subtract_roi(
            ptv_name, oar_combined.Name, "x_PTVm_core", "Red", "Ptv")
        # Perform intersection of PTV with external_shrink
        ptv_core = self._intersection_roi(
            "x_PTVm_core", external_shrink.Name, ptv_core.Name, "Red", "Ptv")
        self._simplify_roi(ptv_core.Name)

        self._delete_roi("x_external_shrink")

        return gtv_core, ptv_core

    def make_ptv_spheres(self, gtv_name, ptv_name):
        hot_ptv_name = "PTVm_6670_" + str(round(self.radius * 2, 1))
        cold_ptv_name = "Eval_PTVm_Avoid_" + str(round(self.radius * 2, 1))
        original_gtv_name = self.params["gtv_name"]

        lattice_params = self.create_sphere_lattice(gtv_name)
        centroid = np.array([lattice_params["target_centroid"].x,
                            lattice_params["target_centroid"].y, lattice_params["target_centroid"].z])

        gtv_struct = Structure(
            {'name': gtv_name, 'case': self.case, 'examination': self.exam})

        optimal_centers, lattice_dict = self._optimize_sphere(
            gtv_struct,
            lattice_params["center_coords"],
            lattice_params["bounding_box"],
            self.xy_spacing,
            self.z_spacing,
            lattice_params["target_centroid"]
        )
        if lattice_dict["num_sphere"] < self.threshold:
            return False

        # Find the coordinate of the hot sphere that is closest to the centroid
        closest_center = optimal_centers[np.argmin(
            np.linalg.norm(optimal_centers - centroid, axis=1))]
        # Set the ICRU point at the coordinate of the closest center
        poi_list = [poi.Name for poi in self.case.PatientModel.PointsOfInterest]
        if "ICRU_X" in poi_list:
            if not hasattr(self.structure_set.PoiGeometries["ICRU_X"].Point, "x"):
                self.structure_set.PoiGeometries["ICRU_X"].Point = {
                    'x': closest_center[0], 'y': closest_center[1], 'z': closest_center[2]}
            else:
                self.structure_set.PoiGeometries["ICRU_X"].Point.x = closest_center[0]
                self.structure_set.PoiGeometries["ICRU_X"].Point.y = closest_center[1]
                self.structure_set.PoiGeometries["ICRU_X"].Point.z = closest_center[2]
        else:
            self.case.PatientModel.CreatePoi(Name="ICRU_X",
                                             Examination=self.exam,
                                             Point={'x': closest_center[0],
                                                    'y': closest_center[1],
                                                    'z': closest_center[2]},
                                             Color='Blue',
                                             Type='DoseRegion')

        cold_centers = np.array(lattice_params["center_coords"])
        cold_centers[:, 0] += 0.5 * self.xy_spacing
        rotation_matrix = np.array([
            [np.cos(np.deg2rad(lattice_dict["angle"])), -
             np.sin(np.deg2rad(lattice_dict["angle"])), 0],
            [np.sin(np.deg2rad(lattice_dict["angle"])), np.cos(
                np.deg2rad(lattice_dict["angle"])), 0],
            [0, 0, 1]
        ])
        cold_centers = np.dot(cold_centers - centroid,
                              rotation_matrix) + centroid
        cold_centers[:, 0] += lattice_dict["shift"][0]
        cold_centers[:, 1] += lattice_dict["shift"][1]
        cold_centers[:, 2] += lattice_dict["shift"][2]

        with CompositeAction("Create Hot Spheres"):
            retval_hot = self._create_roi("Initial_hot_spheres", "Red", "Ptv")

        # Create the spheres
        for idx, center in enumerate(optimal_centers):
            sphere_name = f"H_{idx}"
            sphere = self.create_sphere(
                center, self.radius, sphere_name, "Red")
            # Union the sphere with the hot PTV
            self._union_roi("Initial_hot_spheres", sphere.Name)
            # Delete the sphere
            self.case.PatientModel.RegionsOfInterest[sphere.Name].DeleteRoi()

        # Intersection of hot PTV with original GTVm
        self._intersection_roi("Initial_hot_spheres",
                               original_gtv_name, hot_ptv_name, "Red", "Ptv")

        self.case.PatientModel.RegionsOfInterest["Initial_hot_spheres"].DeleteRoi(
        )

        with CompositeAction("Create Cold Spheres"):
            retval_cold = self._create_roi(
                "Initial_cold_spheres", "Blue", "Ptv")

        max_distance = (2 * (0.5 * self.xy_spacing)**2 +
                        self.z_spacing**2)**0.5 + self.radius * 0.5
        # Create the spheres
        for idx, center in enumerate(cold_centers):
            # We only keep cold spheres that are within max_distance of an optimal sphere
            if np.min(np.linalg.norm(optimal_centers - center, axis=1)) < max_distance:
                sphere_name = f"C_{idx}"
                sphere = self.create_sphere(
                    center, self.radius, sphere_name, "Blue")
                self._union_roi("Initial_cold_spheres", sphere.Name)
                self.case.PatientModel.RegionsOfInterest[sphere.Name].DeleteRoi(
                )

        # Intersection of cold PTV with PTVm_core
        self._intersection_roi("Initial_cold_spheres",
                               ptv_name, cold_ptv_name, "Blue", "Ptv")
        self.case.PatientModel.RegionsOfInterest["Initial_cold_spheres"].DeleteRoi(
        )

        self._simplify_roi(cold_ptv_name)

        return True

    def _inside_box(self, point, bounding_box):
        # Check if the point is inside the bounding box
        if point[0] < bounding_box[0][0] or point[0] > bounding_box[0][1]:
            return False
        if point[1] < bounding_box[1][0] or point[1] > bounding_box[1][1]:
            return False
        if point[2] < bounding_box[2][0] or point[2] > bounding_box[2][1]:
            return False
        return True

    def create_sphere_lattice(self, gtv_name):
        xy_spacing = self.xy_spacing
        z_spacing = self.z_spacing

        # Get the center of GTV
        gtv_obj = self.structure_set.RoiGeometries[gtv_name]
        target_centroid = gtv_obj.GetCenterOfRoi()

        bounding_box = gtv_obj.GetBoundingBox()
        sign = np.sign(target_centroid.x - bounding_box[0].x)
        xp = np.arange(target_centroid.x,
                       bounding_box[1].x + sign * 1.0 * xy_spacing, xy_spacing)
        xn = np.arange(target_centroid.x,
                       bounding_box[0].x - sign * 2.5 * xy_spacing, -xy_spacing)
        x = np.concatenate([xn[::-1], xp[1:]])

        yp = np.arange(target_centroid.y,
                       bounding_box[1].y + sign * 0.5 * xy_spacing, xy_spacing)
        yn = np.arange(target_centroid.y,
                       bounding_box[0].y - sign * 1.5 * xy_spacing, -xy_spacing)
        y = np.concatenate([yn[::-1], yp[1:]])

        x_grid, y_grid = np.meshgrid(x, y)

        # overlay another array of spheres that are staggered by xy_spacing/2
        x1 = x - xy_spacing * 0.5
        y1 = y - xy_spacing * 0.5
        x1_grid, y1_grid = np.meshgrid(x1, y1)

        # flatten the meshgrids
        x_grid = x_grid.flatten()
        y_grid = y_grid.flatten()
        x1_grid = x1_grid.flatten()
        y1_grid = y1_grid.flatten()

        # combine
        x_grid = np.concatenate((x_grid, x1_grid))
        y_grid = np.concatenate((y_grid, y1_grid))

        # find the x and y coordinates of the staggered slice

        x1_grid = x_grid + xy_spacing * 0.5
        y1_grid = y_grid

        z_min = bounding_box[0].z
        z_max = bounding_box[1].z

        # start on the centroid slice
        z = target_centroid.z

        center_coords = []
        while z < z_max + z_spacing:
            for (x, y) in zip(x_grid, y_grid):
                center = [x, y, z]
                center_coords.append(center)
            z += z_spacing
            for (x, y) in zip(x1_grid, y1_grid):
                center = [x, y, z]
                center_coords.append(center)
            z += z_spacing

        # reset z to the centroid slice - z_spacing
        z = target_centroid.z - z_spacing
        while z > z_min - z_spacing * 3:
            for (x, y) in zip(x1_grid, y1_grid):
                center = [x, y, z]
                center_coords.append(center)
            z -= z_spacing
            for (x, y) in zip(x_grid, y_grid):
                center = [x, y, z]
                center_coords.append(center)
            z -= z_spacing

        lattice_params = {
            "center_coords": center_coords,
            "bounding_box": bounding_box,
            "target_centroid": target_centroid,
        }

        return lattice_params

    def create_sphere(self, center, radius, sphere_name, color):
        # Create a sphere at the center with radius
        x, y, z = center
        action_str = "Create"
        with CompositeAction(action_str):
            retval_0 = self._create_roi(sphere_name, color, "Undefined")
            retval_0.CreateSphereGeometry(Radius=radius, Examination=self.exam, Center={
                'x': x, 'y': y, 'z': z}, Representation="TriangleMesh", VoxelSize=None)
        return retval_0

    def _union_roi(self, roi1, roi2):
        # Union two ROIs
        with CompositeAction("Union ROIs"):
            self.case.PatientModel.RegionsOfInterest[roi1].CreateAlgebraGeometry(
                Examination=self.exam, Algorithm="Auto",
                ExpressionA={
                    'Operation': "Union",
                    'SourceRoiNames': [roi1],
                    'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}
                },
                ExpressionB={
                    'Operation': "Union",
                    'SourceRoiNames': [roi2],
                    'MarginSettings': {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0}
                },
                ResultOperation="Union",
                ResultMarginSettings={'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0})

    def _optimize_sphere(self, struct_obj, center_coords, bounding_box, xy_spacing, z_spacing, centroid):
        # We make a dummy coords grid of spacing 1 mm
        coords = {
            "img_pos": [bounding_box[0].x, bounding_box[0].y, min(struct_obj.z_coords)],
            "spacing": [0.1, 0.1, self.slice_spacing]
        }

        x_voxels = np.arange(
            bounding_box[0].x, bounding_box[1].x, coords["spacing"][0])
        y_voxels = np.arange(
            bounding_box[0].y, bounding_box[1].y, coords["spacing"][1])
        z_voxels = np.arange(min(struct_obj.z_coords), max(
            struct_obj.z_coords), coords["spacing"][2])

        coords["num_voxels"] = [len(x_voxels), len(y_voxels), len(z_voxels)]

        x, y = np.meshgrid(x_voxels, y_voxels)
        coords["grid"] = np.array([x, y])
        coords["z_voxels"] = z_voxels
        # We create a mask of the target on the coords grid
        target_mask = struct_obj.get_mask(coords)

        # We want to pad the target mask with 1 row/column of zeros to prevent out of bounds errors
        target_mask = np.pad(target_mask, ((1, 1), (1, 1),
                             (1, 1)), mode='constant', constant_values=0)

        num_shifts = int(xy_spacing * 0.5 * 10) * \
            int(xy_spacing * 10) * int(2 * z_spacing * 10) * 90 // 15
        center_coords = np.array(center_coords)
        center_coords = np.repeat(
            center_coords[np.newaxis, :, :], num_shifts, axis=0)

        centroid = np.array([centroid.x, centroid.y, centroid.z])
        # We create a np array of the shifts
        for i in range(num_shifts):
            shift = self._index_to_shift(i, xy_spacing, z_spacing)
            rotation_matrix = np.array([
                [np.cos(np.deg2rad(shift[3])), -
                 np.sin(np.deg2rad(shift[3])), 0],
                [np.sin(np.deg2rad(shift[3])), np.cos(
                    np.deg2rad(shift[3])), 0],
                [0, 0, 1]
            ])
            center_coords[i, :, :] = np.dot(
                center_coords[i, :, :] - centroid, rotation_matrix) + centroid
            center_coords[i, :, :] += shift[:-1]

        # We add + 1 to the voxels to account for padding
        x_voxels = ((center_coords[:, :, 0] - coords["img_pos"]
                    [0]) / coords["spacing"][0]).astype(int) + 1
        y_voxels = ((center_coords[:, :, 1] - coords["img_pos"]
                    [1]) / coords["spacing"][1]).astype(int) + 1
        z_voxels = ((center_coords[:, :, 2] - coords["img_pos"]
                    [2]) / coords["spacing"][2]).astype(int) + 1

        voxels = (x_voxels, y_voxels, z_voxels)
        np.clip(voxels[0], 0, target_mask.shape[0] - 1, out=voxels[0])
        np.clip(voxels[1], 0, target_mask.shape[1] - 1, out=voxels[1])
        np.clip(voxels[2], 0, target_mask.shape[2] - 1, out=voxels[2])

        masked_voxel = target_mask[voxels]
        num_sphere = np.sum(masked_voxel, axis=1)

        optimal_shift_index = np.argmax(num_sphere)
        optimal_shift = self._index_to_shift(
            optimal_shift_index, xy_spacing, z_spacing)

        optimal_centers = center_coords[optimal_shift_index, :, :]
        optimal_centers = optimal_centers[masked_voxel[optimal_shift_index]]

        lattice_dict = {
            "shift": [float(x) for x in optimal_shift[:-1]],
            "angle": float(optimal_shift[-1]),
            "num_sphere": float(num_sphere[optimal_shift_index])
        }
        print(lattice_dict)
        return optimal_centers, lattice_dict

    def _index_to_shift(self, index, xy_spacing, z_spacing):
        # The shift at index n is given by [x_shift, y_shift, z_shift, angle] where
        x_shift = (index % (xy_spacing * 0.5 * 10)) * 0.1
        y_shift = (index // (xy_spacing * 0.5 * 10)) % (xy_spacing * 10) * 0.1
        z_shift = (index // ((xy_spacing * 10)**2 * 0.5)
                   ) % (2 * z_spacing * 10) * 0.1
        angle = (index // ((xy_spacing * 10)**2 * 0.5 * 2 * z_spacing * 10)) * 15
        return [x_shift, y_shift, z_shift, angle]


if __name__ == '__main__':

    print("Program started...")
    #
    # Check that a patient is open
    # In case no, the script will terminate
    #
    patient = None
    try:
        patient = get_current("Patient")
    except Exception as e:
        MessageBox.Show("No patient loaded.", error_message_header)
        sys.exit()

    #
    # Check that a case is loaded
    # In case no, the script will terminate
    #
    case = None
    try:
        case = get_current("Case")
    except Exception as e:
        MessageBox.Show("No case active.", error_message_header)
        sys.exit()

    if len([exam for exam in case.Examinations]) < 1:
        MessageBox.Show("One or more examinations are required.",
                        error_message_header)
        sys.exit()

    def run_window():
        dialog = MyWindow()

    thread = Thread(ThreadStart(lambda: run_window()))
    thread.SetApartmentState(ApartmentState.STA)
    thread.Start()
    thread.Join()
    print("...program ended")
