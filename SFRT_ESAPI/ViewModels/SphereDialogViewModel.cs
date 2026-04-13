using SFRT_PlanningScript.Models;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Windows;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using ESAPIScript;
using System.Diagnostics;
using System.Threading.Tasks;


namespace SFRT_PlanningScript
{
    public class LatticeParameters
    {
        public float Radius { get; set; }
        public double XShift { get; set; }
        public double YShift { get; set; }
        public float VThresh { get; set; }
        public string BodyId { get; set; }
        public string TargetStructure { get; set; }
        public string PTVLowStructure { get; set; }
        public List<string> OarStructures { get; set; }
        public double BodyMargin { get; set; }
        public double OarMargin { get; set; }
        public bool CouchKick { get; set; }
        public string Energy { get; set; }
        public string MachineId { get; set; }
        public string SphereSize { get; set; }
        public double LatticeSpacing { get; set; }
        public bool HighRes { get; set; }
    }
    public class SphereDialogViewModel : BindableBase
    {
        private LatticeParameters latticeParams;
        private Model _model;
        private EsapiWorker _ew = null;
        private string output;
        public string Output
        {
            get { return output; }
            set { SetProperty(ref output, value); }
        }

        private double xShift;
        public double XShift
        {
            get { return xShift; }
            set { SetProperty(ref xShift, value); }
        }

        private double yShift;
        public double YShift
        {
            get { return yShift; }
            set { SetProperty(ref yShift, value); }
        }

        private float radius;
        public float Radius
        {
            get { return radius; }
            set { SetProperty(ref radius, value); }
        }

        private ObservableCollection<string> oarStructures;

        public ObservableCollection<string> OarStructures
        {
            get { return oarStructures; }
            set { SetProperty(ref oarStructures, value); }
        }
        private List<string> targetStructures;
        public List<string> TargetStructures
        {
            get { return targetStructures; }
            set { SetProperty(ref targetStructures, value); }
        }

        private int selectedOar;
        public int SelectedOar
        {
            get { return selectedOar; }
            set { SetProperty(ref selectedOar, value); }
        }

        private ObservableCollection<string> selectedOars;
        public ObservableCollection<string> SelectedOars
        {
            get { return selectedOars; }
            set
            {
                SetProperty(ref selectedOars, value);
            }
        }

        private int selectedTableOar;
        public int SelectedTableOar
        {
            get { return selectedTableOar; }
            set { SetProperty(ref selectedTableOar, value); }
        }
        private int targetSelected;
        public int TargetSelected
        {
            get { return targetSelected; }
            set { SetProperty(ref targetSelected, value); }
        }

        private int ptvSelected;

        public int PtvSelected
        {
            get { return ptvSelected; }
            set { SetProperty(ref ptvSelected, value); }
        }

        private int bodySelected;
        public int BodySelected
        {
            get { return bodySelected; }
            set
            {
                SetProperty(ref bodySelected, value);
                BodyId = allStructures[value];
            }
        }

        public ObservableCollection<string> EnergyList { get; set; } = new ObservableCollection<string>
        {
            "6X",
            "10X",
            "15X",
            "6X-FFF",
            "10X-FFF"
        };

        public ObservableCollection<string> SphereSize { get; set; } = new ObservableCollection<string>
        {
            "1.5",
            "1.0",
            "0.75",
            "0.5",
        };

        private int selectedSphereSize = 0;

        public int SelectedSphereSize
        {
            get { return selectedSphereSize; }
            set { SetProperty(ref selectedSphereSize, value); }
        }

        private int selectedEnergy = 0;

        public int SelectedEnergy
        {
            get { return selectedEnergy; }
            set { SetProperty(ref selectedEnergy, value); }
        }
        public ObservableCollection<string> MachineIds
        { get; set; } = new ObservableCollection<string>
        {
            "LA16",
            "LA17",
            "LA20",
            "SB_LA_2",
            "ROP_LA_2"
        };

        private int selectedMachineId = 0;
        public int SelectedMachineId
        {
            get { return selectedMachineId; }
            set { SetProperty(ref selectedMachineId, value); }
        }
        private ObservableCollection<string> allStructures;
        public ObservableCollection<string> AllStructures
        {
            get { return allStructures; }
            set { SetProperty(ref allStructures, value); }
        }
        private float vThresh;
        public float VThresh
        {
            get { return vThresh; }
            set { SetProperty(ref vThresh, value); }
        }

        private double LatticeSpacing;

        private bool enableCreate = true;
        public bool EnableCreate
        {
            get { return enableCreate; }
            set { SetProperty(ref enableCreate, value); }
        }

        private bool enableOptimize = true;
        public bool EnableOptimize
        {
            get { return enableOptimize; }
            set { SetProperty(ref enableOptimize, value); }
        }

        private string BodyId;


        private double bodyMargin;
        public double BodyMargin
        {
            get { return bodyMargin; }
            set { SetProperty(ref bodyMargin, value); }
        }

        private double oarMargin;

        public double OarMargin
        {
            get { return oarMargin; }
            set { SetProperty(ref oarMargin, value); }
        }

        private bool couchKick = false;

        public bool CouchKick
        {
            get { return couchKick; }
            set { SetProperty(ref couchKick, value); }
        }

        private bool highRes = false;
        public bool HighRes
        {
            get { return highRes; }
            set { SetProperty(ref highRes, value); }
        }


        public SphereDialogViewModel(EsapiWorker ew = null)
        {
            _ew = ew;
            Initialize();
        }

        private async void Initialize()
        {

            // ctor
            _model = new Model(_ew);
            await _model.InitializeModel();

            // Set UI value defaults
            VThresh = 95;
            // IsHex = true; // default to hex
            XShift = 0;
            YShift = 0;
            Output = " ";

            // Target structures
            targetStructures = new List<string>();
            targetSelected = -1;
            ptvSelected = -1;
            bodySelected = -1;

            bodyMargin = 15.0;
            OarMargin = 15.0;

            selectedOars = new ObservableCollection<string>();
            selectedOar = -1;
            selectedTableOar = -1;

            (OarStructures, AllStructures, TargetStructures, BodyId) = await _model.FetchStructures();
            BodySelected = AllStructures.IndexOf(BodyId);
        }

        private bool PreSpheres()
        {
            Radius = Convert.ToSingle(SphereSize[SelectedSphereSize]) * 10.0f / 2.0f;
            if (Radius == 7.5f)
                LatticeSpacing = 60.0;
            else if (Radius == 5.0f)
                LatticeSpacing = 40.0;
            else
                LatticeSpacing = Radius * 8.0f; // default or fallback
            // Check vol thresh for spheres
            if (VThresh > 100 || VThresh < 0)
            {
                MessageBox.Show("Volume threshold must be between 0 and 100");
                return false;
            }

            // Check target
            if (targetSelected == -1)
            {
                MessageBox.Show("Must have target selected, cancelling operation.");
                return false;
            }

            if (Radius <= 0)
            {
                MessageBox.Show("Radius must be greater than zero.");
                return false;
            }

            if (LatticeSpacing < Radius * 2)
            {
                var buttons = MessageBoxButton.OKCancel;
                var result = MessageBox.Show($"WARNING: Sphere center spacing is less than sphere diameter ({Radius * 2}) mm.\n Continue?", "", buttons);
                return result == MessageBoxResult.OK;
            }

            // Check that "BODY" structure exists
            if (bodySelected == -1)
            {
                MessageBox.Show("Please select a body structure.");
                return false;
            }

            return true;
        }

        public async Task CreateLattice()
        {

            if (!PreSpheres())
            {
                return;
            }
            latticeParams = new LatticeParameters()
            {
                Radius = Radius,
                XShift = XShift,
                YShift = YShift,
                VThresh = VThresh,
                BodyId = BodyId,
                TargetStructure = TargetStructures[TargetSelected],
                PTVLowStructure = TargetStructures[PtvSelected],
                OarStructures = SelectedOars.ToList(),
                BodyMargin = BodyMargin,
                OarMargin = OarMargin,
                CouchKick = CouchKick,
                Energy = EnergyList[SelectedEnergy],
                MachineId = MachineIds[SelectedMachineId],
                SphereSize = SphereSize[SelectedSphereSize],
                LatticeSpacing = LatticeSpacing,
                HighRes = HighRes
            };

            EnableCreate = false;

            var progress = new Progress<string>(message =>
            {
                Output += message + "\n";
            });
            Output += "Lattice creation in progress... This will take 5 to 10 minutes.";
            await _model.BuildSpheres(latticeParams, true, true, progress);
            Output += "Lattice creation complete. Setting up beams...\n";
            await _model.SetupBeams(latticeParams);
            Output += "Beams setup complete. Review and proceed to optimization.\n";
            MessageBox.Show("Script execution complete.");
        }

        public async Task Optimize()
        {
            latticeParams = new LatticeParameters()
            {
                Radius = Radius,
                XShift = XShift,
                YShift = YShift,
                VThresh = VThresh,
                BodyId = BodyId,
                TargetStructure = TargetStructures[TargetSelected],
                PTVLowStructure = TargetStructures[PtvSelected],
                OarStructures = SelectedOars.ToList(),
                BodyMargin = BodyMargin,
                OarMargin = OarMargin,
                CouchKick = CouchKick,
                Energy = EnergyList[SelectedEnergy],
                MachineId = MachineIds[SelectedMachineId],
                SphereSize = SphereSize[SelectedSphereSize],
                HighRes = HighRes
            };

            EnableOptimize = false;
            var progress = new Progress<string>(message =>
            {
                Output += message + "\n";
            });
            _model.LoadDefaultOptimizationParameters();
            Output += "Starting optimization...\n";
            await _model.OptimizeLattice(progress, latticeParams);
            Output += "Optimization complete.\n";
            MessageBox.Show("Optimization complete.");
        }

        public void AddRoi()
        {
            // Add selected roi to list
            if (SelectedOar != -1)
            {
                SelectedOars.Add(OarStructures[SelectedOar]);
                OarStructures.RemoveAt(SelectedOar);
                SelectedOar = -1;
            }
        }
        public void RemoveRoi()
        {
            // Remove selected roi from list
            if (SelectedTableOar != -1)
            {
                OarStructures.Add(SelectedOars[SelectedTableOar]);
                SelectedOars.RemoveAt(SelectedTableOar);
                SelectedTableOar = SelectedTableOar - 1;
            }
        }

    }
}
