using SFRT_PlanningScript.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using System.Text.RegularExpressions;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;


namespace SFRT_PlanningScript.ViewModels
{
    public class GridMLCDialogViewModel : INotifyPropertyChanged
    {

        bool shrink_by_1 = true;

        public bool ShrinkBy1
        {
            get { return shrink_by_1; }
            set { shrink_by_1 = value; OnPropertyChanged(); }
        }
        float offset = 0.5f;

        public float Offset
        {
            get { return offset; }
            set { offset = value; OnPropertyChanged(); }
        }
        float squareSize;

        public float SquareSize
        {
            get { return squareSize; }
            set { squareSize = value; OnPropertyChanged(); }
        }
        double canvasHeight;
        public double CanvasHeight
        {
            get { return canvasHeight; }
            set { canvasHeight = value; OnPropertyChanged(); }
        }
        double canvasWidth;
        public double CanvasWidth
        {
            get { return canvasWidth; }
            set { canvasWidth = value; OnPropertyChanged(); }
        }

        double bbXs;
        double bbYs;
        double bbXe;
        double bbYe;

        private CoordinateConverter xConv;
        private CoordinateConverter yConv;

        public ScriptContext context;

        public GridMLCDialogViewModel(ScriptContext currentContext)
        {
            context = currentContext;

            //ui 'consts'
            canvasHeight = 300;
            canvasWidth = 400;

            //ui defaults
            squareSize = 10;
            shrink_by_1 = true;
            offset = 0.5f;
            //hidden defaults
            bbXs = 50;
            bbXe = 250;
            bbYs = 50;
            bbYe = 200;


        }
        private void UpdateCanvasScaling()
        {
            double fullWidth = bbXe - bbXs;
            double fullHeight = bbYe - bbYs;

            double largerScaler = fullWidth / canvasWidth > fullHeight / canvasHeight ? fullWidth / canvasWidth : fullHeight / canvasHeight;
            double widthMargin = 0.5 * (largerScaler - fullWidth / canvasWidth);
            double heightMargin = 0.5 * (largerScaler - fullHeight / canvasHeight);

            xConv = new CoordinateConverter(bbXs - widthMargin * canvasWidth, bbXe + widthMargin * canvasWidth, canvasWidth);

            yConv = new CoordinateConverter(bbYs - heightMargin * canvasHeight, bbYe + heightMargin * canvasHeight, canvasHeight);
        }
        private float get_leaf_width(int idx)
        {
            if (idx < 10 || idx > 49)
                return 10.0f;
            else
                return 5.0f;
        }

        private void FindSquareCenters(float[,] bankA, float[,] bankB, List<List<float[]>> squareCenters_1, List<List<float[]>> squareCenters_2)
        {
            var firstOpen = 60;
            var lastOpen = 0;
            // Find first and last open leaves
            // We consider leaf pairs that are separated by less than a square size
            // To be closed
            for (int i = 0; i < 60; i++)
            {
                var leafDiff = Math.Abs(bankA[i, 0] - bankB[i, 0]);
                if (leafDiff > squareSize)
                {
                    firstOpen = i;
                    break;
                }
            }
            for (int i = 59; i >= 0; i--)
            {
                var leafDiff = Math.Abs(bankA[i, 0] - bankB[i, 0]);
                if (leafDiff > squareSize)
                {
                    lastOpen = i;
                    break;
                }
            }

            var leftmost_x = 200f;
            var rightmost_x = -200f;
            for (int i = firstOpen; i < lastOpen; i++)
            {
                if (bankA[i, 0] < leftmost_x)
                {
                    leftmost_x = bankA[i, 0];
                }
                if (bankB[i, 0] > rightmost_x)
                {
                    rightmost_x = bankB[i, 0];
                }
            }
            if (shrink_by_1)
            {
                firstOpen += 1;
                lastOpen -= 1;
            }
            var currLeaf = firstOpen;

            var rowCenters_1 = new List<float[]>();
            var center = new float[2];

            // Put the first center square at the leftmost_x + squareSize
            center[0] = leftmost_x + squareSize * offset;
            center[1] = bankA[currLeaf, 1] - get_leaf_width(currLeaf) / 2 + squareSize / 2;
            // MessageBox.Show($"Center: {center[0]}, {center[1]}");

            rowCenters_1.Add(new float[2] { center[0], center[1] });

            var highest_y = bankA[lastOpen, 1];

            // First Row;
            while (center[0] + 4 * squareSize < rightmost_x)
            {
                center[0] = center[0] + 4 * squareSize;
                center[1] = center[1];
                rowCenters_1.Add(new float[2] { center[0], center[1] });
            }

            squareCenters_1.Add(rowCenters_1);

            var current_y = new float();
            current_y = rowCenters_1[0][1] + 2 * squareSize;
            // While y < highest_y, copy row and shift y
            while (current_y < highest_y)
            {
                var rowCenters = new List<float[]>();
                foreach (var square in rowCenters_1)
                {
                    center[0] = square[0];
                    center[1] = current_y;
                    rowCenters.Add(new float[2] { center[0], center[1] });
                }
                squareCenters_1.Add(rowCenters);
                current_y += 2 * squareSize;
            }

            // Copy squareCenters_1 but shift everything by 1 square size in x and 2 square sizes in y
            foreach (var row in squareCenters_1)
            {
                var rowCenters = new List<float[]>();
                foreach (var square in row)
                {
                    center[0] = square[0] + 2 * squareSize;
                    center[1] = square[1] + squareSize;
                    rowCenters.Add(new float[2] { center[0], center[1] });
                }
                squareCenters_2.Add(rowCenters);
            }

        }

        private float Constrain_Inside(float x, float min, float max)
        {
            return Math.Max(min, Math.Min(max, x));
        }


        public void CreateGrid()
        {
            //Start prepare the patient
            context.Patient.BeginModifications();

            // Grab the first field MLC position
            var firstBeam = context.PlanSetup.Beams.First();
            var leafPositions = firstBeam.ControlPoints[0].LeafPositions;

            // Convert the leaf positions to a list of points
            // x = leafPositions, y = leaf_index * leaf_width
            var bankA = new float[60, 2];
            var bankB = new float[60, 2];


            bankA[0, 0] = leafPositions[0, 0];
            bankA[0, 1] = get_leaf_width(0) * 0.5f;
            bankB[0, 0] = leafPositions[1, 0];
            bankB[0, 1] = get_leaf_width(0) * 0.5f;

            for (int i = 1; i < 60; i++)
            {
                bankA[i, 0] = leafPositions[0, i];
                bankA[i, 1] = get_leaf_width(i - 1) * 0.5f + get_leaf_width(i) * 0.5f + bankA[i - 1, 1];
                bankB[i, 0] = leafPositions[1, i];
                bankB[i, 1] = get_leaf_width(i - 1) * 0.5f + get_leaf_width(i) * 0.5f + bankB[i - 1, 1];
            }

            var firstOpen = 60;
            var lastOpen = 0;
            // Find first and last open leaves
            // We consider leaf pairs that are separated by less than a square size
            // To be closed
            for (int i = 0; i < 60; i++)
            {
                var leafDiff = Math.Abs(bankA[i, 0] - bankB[i, 0]);
                if (leafDiff > squareSize)
                {
                    firstOpen = i;
                    break;
                }
            }
            for (int i = 59; i >= 0; i--)
            {
                var leafDiff = Math.Abs(bankA[i, 0] - bankB[i, 0]);
                if (leafDiff > squareSize)
                {
                    lastOpen = i;
                    break;
                }
            }

            var squareCenters_1 = new List<List<float[]>>();
            var squareCenters_2 = new List<List<float[]>>();
            FindSquareCenters(bankA, bankB, squareCenters_1, squareCenters_2);

            var newApertures = new List<List<float[,]>>();

            for (int j = 0; j < squareCenters_1[0].Count; j++)
            {
                // Each column of squareCenters_1 and 2 are responsible for a single aperture
                var new_bankA = new float[60, 2];
                var new_bankB = new float[60, 2];

                Array.Copy(bankA, new_bankA, 60 * 2);
                Array.Copy(bankB, new_bankB, 60 * 2);

                var all_leaf_involved = new HashSet<int>();
                for (int i = 0; i < squareCenters_1.Count; i++)
                {
                    var square1 = squareCenters_1[i][j];
                    // MessageBox.Show($"Square1: {square1[0]}, {square1[1]}");
                    var leaf_involved = new List<int>();
                    // Any leaf that is within squareSize / 2 of the square center x
                    for (int k = firstOpen; k < lastOpen; k++)
                    {
                        if (Math.Abs(bankA[k, 1] - square1[1]) < squareSize / 2 - 0.001f)
                        {
                            // MessageBox.Show($"Leaf Involved: {k}");
                            // MessageBox.Show($"BankA: {bankA[k, 1]}, Square1: {square1[1]}");
                            leaf_involved.Add(k);
                            all_leaf_involved.Add(k);
                        }
                    }
                    for (int k = 0; k < leaf_involved.Count; k++)
                    {
                        // MessageBox.Show($"Leaf Involved: {leaf_involved[k]}");
                        new_bankA[leaf_involved[k], 0] = square1[0] - squareSize / 2;
                        new_bankB[leaf_involved[k], 0] = square1[0] + squareSize / 2;

                        new_bankA[leaf_involved[k], 0] = Constrain_Inside(new_bankA[leaf_involved[k], 0], bankA[leaf_involved[k], 0], bankB[leaf_involved[k], 0]);
                        new_bankB[leaf_involved[k], 0] = Constrain_Inside(new_bankB[leaf_involved[k], 0], bankA[leaf_involved[k], 0], bankB[leaf_involved[k], 0]);
                        // MessageBox.Show($"New BankA: {new_bankA[leaf_involved[k], 0]}, New BankB: {new_bankB[leaf_involved[k], 0]}");
                    }

                    var square2 = squareCenters_2[i][j];
                    // MessageBox.Show($"Square2: {square2[0]}, {square2[1]}");
                    leaf_involved = new List<int>();

                    for (int k = firstOpen; k < lastOpen; k++)
                    {
                        if (Math.Abs(bankA[k, 1] - square2[1]) <= squareSize / 2)
                        {
                            leaf_involved.Add(k);
                            all_leaf_involved.Add(k);
                        }
                    }
                    for (int k = 0; k < leaf_involved.Count; k++)
                    {
                        new_bankA[leaf_involved[k], 0] = square2[0] - squareSize / 2;
                        new_bankB[leaf_involved[k], 0] = square2[0] + squareSize / 2;

                        new_bankA[leaf_involved[k], 0] = Constrain_Inside(new_bankA[leaf_involved[k], 0], bankA[leaf_involved[k], 0], bankB[leaf_involved[k], 0]);
                        new_bankB[leaf_involved[k], 0] = Constrain_Inside(new_bankB[leaf_involved[k], 0], bankA[leaf_involved[k], 0], bankB[leaf_involved[k], 0]);
                    }
                }

                // Loop from first to last, if not in all_leaf_involved, close
                var closed_involved = 0;
                for (int i = firstOpen; i <= lastOpen; i++)
                {
                    if (!all_leaf_involved.Contains(i))
                    {
                        new_bankA[i, 0] = bankA[0, 0];
                        new_bankB[i, 0] = bankB[0, 0];
                    }
                    else
                    {
                        if (new_bankA[i, 0] == new_bankB[i, 0])
                        {
                            new_bankA[i, 0] = bankA[0, 0];
                            new_bankB[i, 0] = bankB[0, 0];
                            closed_involved++;
                        }
                    }
                }
                if (closed_involved < all_leaf_involved.Count)
                {
                    // var print_str = "";
                    // // Print leaf positions
                    // for (int i = 0; i < 60; i++)
                    // {
                    //     print_str += $"({new_bankA[i, 0]}, {new_bankB[i, 0]}), {i};";
                    // }
                    // MessageBox.Show(print_str);

                    var newAperture = new List<float[,]> { new_bankA, new_bankB };
                    newApertures.Add(newAperture);
                }
                // throw new Exception("stop");

            }

            string energy = firstBeam.EnergyModeDisplayName;
            string fluence = null;
            Match EMode = Regex.Match(energy, @"^([0-9]+[A-Z]+)-?([A-Z]+)?", RegexOptions.IgnoreCase);  //format is... e.g. 6X(-FFF)
            if (EMode.Success)
            {
                if (EMode.Groups[2].Length > 0)  // fluence mode
                {
                    energy = EMode.Groups[1].Value;
                    fluence = EMode.Groups[2].Value;
                } // else normal modes uses default in decleration
            }
            ExternalBeamMachineParameters machineParameters = new ExternalBeamMachineParameters(firstBeam.TreatmentUnit.Id, energy, firstBeam.DoseRate, firstBeam.Technique.Id, fluence);
            // 
            var jawPositions = firstBeam.ControlPoints[0].JawPositions;
            var collimatorAngle = firstBeam.ControlPoints[0].CollimatorAngle;
            var gantryAngle = firstBeam.ControlPoints[0].GantryAngle;
            var patientSupportAngle = firstBeam.ControlPoints[0].PatientSupportAngle;
            VVector isocenter = firstBeam.IsocenterPosition;

            foreach (var aperture in newApertures)
            {
                try
                {
                    // Create a new beam with the new apertures
                    // Convert bankA and bankB to leaf positions
                    var leafPositions_temp = new float[2, 60];
                    for (int i = 0; i < 60; i++)
                    {
                        leafPositions_temp[0, i] = aperture[0][i, 0];
                        leafPositions_temp[1, i] = aperture[1][i, 0];
                    }
                    var newBeam = context.ExternalPlanSetup.AddMLCBeam(machineParameters, leafPositions_temp, jawPositions, collimatorAngle, gantryAngle, patientSupportAngle, isocenter);
                    var editableParams = newBeam.GetEditableParameters();

                    editableParams.ControlPoints.First().LeafPositions = leafPositions_temp;
                    editableParams.WeightFactor = firstBeam.WeightFactor;
                    newBeam.ApplyParameters(editableParams);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    // Print leaf positions
                    var print_str = "";
                    for (int i = 0; i < 60; i++)
                    {
                        print_str += $"({aperture[0][i, 0]}, {aperture[1][i, 0]}) ;";
                    }
                    MessageBox.Show(print_str);
                    break;
                }
            }
            // Exit
            MessageBox.Show("Grid MLCs created successfully, please close the script");
        }

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion INotifyPropertyChanged
    }
}
