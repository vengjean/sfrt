using ESAPIScript;
using Prism.Mvvm;
using SFRT_PlanningScript.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using VMS.TPS.Common.Model;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using static SFRT_PlanningScript.SphereDialogViewModel;


namespace SFRT_PlanningScript
{
    public class Model
    {
        private double mask_spacing = 1.0;
        private double resolution = 1.0;
        private EsapiWorker _ew;

        bool HighRes = false;
        private float Radius;
        private double LatticeSpacing;

        private double bodyMargin;

        private string BodyId;

        private OptimizationSetup Optimizer;

        private double OarMargin;

        private StructureSet _ss;
        public Model(EsapiWorker ew)
        {
            _ew = ew;
        }

        public async Task<bool> InitializeModel()
        {
            await _ew.AsyncRunPlanContext((pat, ps) =>
            {
                pat.BeginModifications();
                _ss = ps.StructureSet;
            });
            return true;
        }

        public async Task<(ObservableCollection<string>, ObservableCollection<string>, List<string>, string)> FetchStructures()
        {
            ObservableCollection<string> oarStructures = new ObservableCollection<string>();
            ObservableCollection<string> allStructures = new ObservableCollection<string>();
            List<string> targetStructures = new List<string>();
            string defaultBodyId = "";

            await _ew.AsyncRunPlanContext((pat, ps) =>
            {
                var ss = ps.StructureSet;

                foreach (var i in ss.Structures)
                {
                    allStructures.Add(i.Id);
                    // if dicom type is not PTV, GTV or CTV, skip
                    oarStructures.Add(i.Id);
                    if (i.DicomType != "PTV" && i.DicomType != "GTV" && i.DicomType != "CTV" && !i.Id.Contains("PTV"))
                    {
                        continue;
                    }
                    targetStructures.Add(i.Id);
                }

                // Search for any structure with "BODY" in the name or "Body" in the name
                for (int idx = 0; idx < allStructures.Count; idx++)
                {
                    if (allStructures[idx].ToLower().Contains("body"))
                    {
                        defaultBodyId = allStructures[idx];
                        break;
                    }
                }
                try
                {
                    // remove the body structure from the list of OARs
                    oarStructures.Remove(defaultBodyId);
                }
                catch (Exception)
                {
                }
            });

            return (oarStructures, allStructures, targetStructures, defaultBodyId);
        }

        private Structure CreateStructure(StructureSet structureSet, string structName, bool showMessage, bool makeHiRes, string structType = "PTV")
        {
            string msg = $"New structure ({structName}) created.";
            var prevStruct = structureSet.Structures.FirstOrDefault(x => x.Id == structName);
            if (prevStruct != null)
            {
                structureSet.RemoveStructure(prevStruct);
                msg += " Old structure overwritten.";
            }

            var structure = structureSet.AddStructure(structType, structName);

            // TEMPORARY -> Need to bring it back to highres

            if (makeHiRes)
            {
                structure.ConvertToHighResolution();
                msg += " Converted to Hi-Res";
            }

            if (showMessage) { MessageBox.Show(msg); }
            return structure;
        }

        private void AddContoursToMain(StructureSet structureSet, ref Structure PrimaryStructure, ref Structure SecondaryStructure)
        {
            // Loop through each image plane
            // { foreach (var segment in contours) { lowResSSource.AddContourOnImagePlane(segment, j); } }
            for (int z = 0; z < structureSet.Image.ZSize; ++z)
            {
                var contours = SecondaryStructure.GetContoursOnImagePlane(z);
                foreach (var seg in contours)
                {
                    PrimaryStructure.AddContourOnImagePlane(seg, z);
                }
            }
        }

        private void BuildSphere(Structure parentStruct, VVector center, float r)//, Structure secondStructure = null)
        {
            double z_center = center.z;
            double min_z = z_center - r;
            // Find the closest slice number to the minimum z value
            int min_z_idx = (int)Math.Floor((min_z - _ss.Image.Origin.z) / _ss.Image.ZRes);
            // Find the closest slice number to the maximum z value
            int max_z_idx = (int)Math.Ceiling((z_center + r - _ss.Image.Origin.z) / _ss.Image.ZRes);

            // Make sure they are above 0 and below the max number of slices
            min_z_idx = Math.Max(min_z_idx, 0);
            max_z_idx = Math.Min(max_z_idx, _ss.Image.ZSize);
            if (min_z_idx == 0 || max_z_idx == _ss.Image.ZSize)
            {
                MessageBox.Show("Sphere is out of image bounds - ERROR");
            }

            for (int z = min_z_idx; z < max_z_idx; ++z)
            {
                double zCoord = z * (_ss.Image.ZRes) + _ss.Image.Origin.z;

                // For each slice find in plane radius
                var z_diff = Math.Abs(zCoord - center.z);
                if (z_diff > r) // If we are out of range of the sphere continue
                {
                    continue;
                }

                // Otherwise make spheres
                var r_z = Math.Sqrt(Math.Pow(r, 2.0) - Math.Pow(z_diff, 2.0));
                var contour = CreateContour(center, r_z, 15);
                parentStruct.AddContourOnImagePlane(contour, z);
            }
        }

        private List<double> Arange(double start, double stop, double step)
        {
            //log.Debug($"Arange with start stop step = {start} {stop} {step}\n");
            var retval = new List<double>();
            var currentval = start;
            while (currentval < stop)
            {
                retval.Add(currentval);
                currentval += step;
            }
            return retval;
        }

        private List<VVector> BuildGrid(List<double> xcoords, List<double> ycoords, List<double> zcoords)
        {
            var retval = new List<VVector>();
            foreach (var x in xcoords)
            {
                foreach (var y in ycoords)
                {
                    foreach (var z in zcoords)
                    {
                        var pt = new VVector(x, y, z);

                        retval.Add(pt);
                    }
                }
            }

            return retval;
        }

        private List<VVector> BuildHexGrid(double Xstart, double Xsize, double Ystart, double Ysize, double Zstart, double Zsize)
        {
            double A = LatticeSpacing;
            var retval = new List<VVector>();

            void CreateLayer(double zCoord, double x0, double y0)
            {
                // create planar hexagonal sphere packing grid
                var yeven = Arange(y0, y0 + Ysize, A);
                var xeven = Arange(x0, x0 + Xsize, A);
                foreach (var y in yeven)
                {
                    foreach (var x in xeven)
                    {
                        retval.Add(new VVector(x, y, zCoord));
                        retval.Add(new VVector(x + (A / 2.0), y + (A / 2.0), zCoord));
                    }
                }
            }

            foreach (var z in Arange(Zstart, Zstart + Zsize, A))
            {
                CreateLayer(z, Xstart, Ystart);
                CreateLayer(z + (A / 2.0), Xstart + (A / 2.0), Ystart);

            }

            return retval;
        }

        private List<VVector> CheckGridInContour(List<VVector> grid, Structure target)
        {
            List<VVector> correctGrid = new List<VVector>();
            for (int i = 0; i < grid.Count; i++)
            {
                if (target.IsPointInsideSegment(grid[i]))
                {
                    correctGrid.Add(grid[i]);
                }
            }
            return correctGrid;
        }

        private VVector index_to_shift(int idx)
        {
            var x_shift = (idx % (LatticeSpacing * 0.5)) * resolution;
            var y_shift = (Math.Floor(idx / (LatticeSpacing * 0.5)) % (LatticeSpacing)) * resolution;
            var z_shift = (Math.Floor(idx / (LatticeSpacing * 0.5 * LatticeSpacing)) % (LatticeSpacing)) * resolution;
            return new VVector(x_shift, y_shift, z_shift);
        }

        private List<VVector> CheckColdGrid(List<VVector> cold_grid, List<VVector> grid)
        {
            List<VVector> correctGrid = new List<VVector>();
            double max_distance = Math.Sqrt(3 * Math.Pow(0.5 * LatticeSpacing, 2)) + Radius * 0.5;
            for (int i = 0; i < cold_grid.Count; i++)
            {
                for (int j = 0; j < grid.Count; j++)
                {
                    if (VVector.Distance(grid[j], cold_grid[i]) < max_distance)
                    {
                        correctGrid.Add(cold_grid[i]);
                        break;
                    }
                }
                // Find at least one point within grid that is within 2*spacing of cold grid point
            }
            return correctGrid;
        }

        private List<VVector> CheckTemplateGrid(List<VVector> grid_template, List<VVector> grid)
        {
            List<VVector> correctGrid = new List<VVector>();
            double max_distance = Math.Sqrt(3 * Math.Pow(0.5 * LatticeSpacing, 2)) + Radius * 0.5;
            for (int i = 0; i < grid_template.Count; i++)
            {
                // Check if grid_template[i] is contained in grid
                if (grid.Contains(grid_template[i]))
                {
                    continue;
                }
                for (int j = 0; j < grid.Count; j++)
                {
                    if (VVector.Distance(grid[j], grid_template[i]) < max_distance)
                    {
                        correctGrid.Add(grid_template[i]);
                        break;
                    }
                }
                // Find at least one point within grid that is within 2*spacing of cold grid point
            }
            return correctGrid;
        }

        public void CreatePrv(List<string> SelectedOars)
        {
            var body = _ss.Structures.FirstOrDefault(x => x.Id == BodyId);

            // Create a new structure
            // var prv = _ss.AddStructure("Control", "zzz_SFRT_PRV");
            CreateStructure(_ss, "zzz_SFRT_PRV", false, false, "Control");
            var prv = _ss.Structures.FirstOrDefault(x => x.Id == "zzz_SFRT_PRV");
            // Add contours to the new structure
            // AddContoursToMain(_ss, ref prv, ref body);
            // crop body_margin mm from body
            prv.SegmentVolume = body.SegmentVolume.Margin(-bodyMargin);

            // Loop through SelectedOars and subtract contours to the new structure with margin
            foreach (var oar in SelectedOars)
            {
                var oar_contour = _ss.Structures.FirstOrDefault(x => x.Id == oar);
                // Check if oar_contour has contours
                if (!oar_contour.HasSegment)
                {
                    continue;
                }

                if (oar_contour.IsHighResolution)
                {
                    CreateStructure(_ss, "zzz_temp_prv", false, false, "Control");
                    var temp_prv = _ss.Structures.FirstOrDefault(x => x.Id == "zzz_temp_prv");
                    AddContoursToMain(_ss, ref temp_prv, ref oar_contour);
                    prv.SegmentVolume = prv.SegmentVolume.Sub(temp_prv.SegmentVolume.Margin(OarMargin));
                    _ss.RemoveStructure(temp_prv);
                }
                else
                {
                    prv.SegmentVolume = prv.SegmentVolume.Sub(oar_contour.SegmentVolume.Margin(OarMargin));
                }
            }

            // do prv = body - prv
            prv.SegmentVolume = body.SegmentVolume.Sub(prv.SegmentVolume);

        }

        private bool[,,] GenerateTargetMask(Structure target)
        {
            var bounds = target.MeshGeometry.Bounds;
            double min_x = bounds.X;
            double max_x = bounds.X + bounds.SizeX;

            double min_y = bounds.Y;
            double max_y = bounds.Y + bounds.SizeY;

            double min_z = bounds.Z;
            double max_z = bounds.Z + bounds.SizeZ;


            // Create a 3d boolean array spanning 3 width of the target bounding box
            int num_voxels_x = Convert.ToInt32(Math.Abs(max_x - min_x) / mask_spacing) + 1;
            int num_voxels_y = Convert.ToInt32(Math.Abs(max_y - min_y) / mask_spacing) + 1;
            int num_voxels_z = Convert.ToInt32(Math.Abs(max_z - min_z) / mask_spacing) + 1;

            bool[,,] target_mask = new bool[num_voxels_x, num_voxels_y, num_voxels_z];

            for (int i = 0; i < num_voxels_x; i++)
            {
                for (int j = 0; j < num_voxels_y; j++)
                {
                    for (int k = 0; k < num_voxels_z; k++)
                    {
                        var pt = new VVector(min_x + i * mask_spacing, min_y + j * mask_spacing, min_z + k * mask_spacing);
                        target_mask[i, j, k] = target.IsPointInsideSegment(pt);
                    }
                }
            }
            return target_mask;
        }

        private List<VVector> CheckGridInMask(List<VVector> grid, bool[,,] mask, Structure target)
        {
            var bounds = target.MeshGeometry.Bounds;
            double min_x = bounds.X - bounds.SizeX;
            double max_x = bounds.X + 2 * bounds.SizeX;
            double min_y = bounds.Y - bounds.SizeY;
            double max_y = bounds.Y + 2 * bounds.SizeY;
            double min_z = bounds.Z - bounds.SizeZ;
            double max_z = bounds.Z + 2 * bounds.SizeZ;
            // Spacing is 1 mm

            List<VVector> correctGrid = new List<VVector>();
            foreach (var pt in grid)
            {
                if (pt.x < min_x || pt.x > max_x || pt.y < min_y || pt.y > max_y || pt.z < min_z || pt.z > max_z)
                {
                    continue;
                }
                int x_idx = Convert.ToInt32(Math.Floor(pt.x - min_x));
                int y_idx = Convert.ToInt32(Math.Floor(pt.y - min_y));
                int z_idx = Convert.ToInt32(Math.Floor(pt.z - min_z));

                if (mask[x_idx, y_idx, z_idx])
                {
                    correctGrid.Add(pt);
                }
            }
            return correctGrid;
        }

        private List<int> SearchGrid(List<VVector> grid, bool[,,] mask, Structure target)
        {
            int max_num_shifts = Convert.ToInt32(Math.Round(LatticeSpacing * 0.5 * LatticeSpacing * LatticeSpacing / resolution));
            // int max_grid_per_shift = grid.Count / max_num_shifts;

            var bounds = target.MeshGeometry.Bounds;
            double min_x = bounds.X;
            double max_x = bounds.X + bounds.SizeX;
            double min_y = bounds.Y;
            double max_y = bounds.Y + bounds.SizeY;
            double min_z = bounds.Z;
            double max_z = bounds.Z + bounds.SizeZ;

            int[] shift_count = new int[max_num_shifts];

            // var temp_grid = new List<VVector>();

            // // sum all good_grid in increments of max
            // for (int shift_idx = 0; shift_idx < max_num_shifts; shift_idx++)
            // {
            //     var count = 0;
            //     var shift = index_to_shift(shift_idx);
            //     temp_grid = grid.Select(x => new VVector(x.x + shift.x, x.y + shift.y, x.z + shift.z)).ToList();
            //     for (int i = 0; i < temp_grid.Count; i++)
            //     {
            //         int x_idx = Convert.ToInt32(Math.Floor(temp_grid[i].x - min_x));
            //         int y_idx = Convert.ToInt32(Math.Floor(temp_grid[i].y - min_y));
            //         int z_idx = Convert.ToInt32(Math.Floor(temp_grid[i].z - min_z));

            //         try
            //         {
            //             if (mask[x_idx, y_idx, z_idx])
            //             {
            //                 count++;
            //             }
            //         }
            //         catch (Exception)
            //         {
            //         }
            //     }

            //     shift_count[shift_idx] = count;
            // }
            var shifted_grids = new List<List<VVector>>(max_num_shifts);

            // Precompute all shifted grids
            for (int shift_idx = 0; shift_idx < max_num_shifts; shift_idx++)
            {
                var shift = index_to_shift(shift_idx);
                var temp_grid = grid.Select(x => new VVector(x.x + shift.x, x.y + shift.y, x.z + shift.z)).ToList();
                shifted_grids.Add(temp_grid);
            }

            // Sum all good_grid in increments of max
            for (int shift_idx = 0; shift_idx < max_num_shifts; shift_idx++)
            {
                var count = 0;
                var temp_grid = shifted_grids[shift_idx];
                for (int i = 0; i < temp_grid.Count; i++)
                {
                    int x_idx = Convert.ToInt32(Math.Floor(temp_grid[i].x - min_x) / mask_spacing);
                    int y_idx = Convert.ToInt32(Math.Floor(temp_grid[i].y - min_y) / mask_spacing);
                    int z_idx = Convert.ToInt32(Math.Floor(temp_grid[i].z - min_z) / mask_spacing);

                    if (x_idx >= 0 && x_idx < mask.GetLength(0) &&
                        y_idx >= 0 && y_idx < mask.GetLength(1) &&
                        z_idx >= 0 && z_idx < mask.GetLength(2) &&
                        mask[x_idx, y_idx, z_idx])
                    {
                        count++;
                    }
                }
                shift_count[shift_idx] = count;
            }


            // find max count shift
            List<int> max_index_list = new List<int>();
            int max_count = 0;

            for (int i = 0; i < max_num_shifts; i++)
            {
                if (shift_count[i] > max_count)
                {
                    max_count = shift_count[i];
                    max_index_list = new List<int> { i };
                }
                else if (shift_count[i] == max_count)
                {
                    max_index_list.Add(i);
                }
            }

            return max_index_list;
        }

        public async Task BuildSpheres(LatticeParameters latticeParams, bool makeIndividual, bool alignGrid, IProgress<string> progress = null)
        {
            await _ew.AsyncRunPlanContext((pat, ps) =>
            {

                StructureSet structureSet = ps.StructureSet;

                BodyId = latticeParams.BodyId;
                Radius = Convert.ToSingle(latticeParams.SphereSize) * 10.0f / 2.0f;
                bodyMargin = latticeParams.BodyMargin;
                OarMargin = latticeParams.OarMargin;
                LatticeSpacing = latticeParams.LatticeSpacing;
                HighRes = latticeParams.HighRes;

                // Start timer
                var sw = new Stopwatch();
                sw.Start();

                // Total lattice structure with all spheres
                Structure structMain = null;
                Structure structMain_cold = null;

                var target_name = latticeParams.TargetStructure;
                var target_initial = structureSet.Structures.Where(x => x.Id == target_name).First();
                Structure target = null;
                bool deleteAutoTarget = false;
                Structure target_initial_temp = null;

                var ptv_low_name = latticeParams.PTVLowStructure;
                var ptv_low = structureSet.Structures.FirstOrDefault(x => x.Id == ptv_low_name);
                progress?.Report("\nPreparing structures ...");

                CreateStructure(structureSet, "zzz_GTV_core", false, false);
                target = structureSet.Structures.FirstOrDefault(x => x.Id == "zzz_GTV_core");
                AddContoursToMain(structureSet, ref target, ref target_initial);
                if (target == null)
                {
                    return;
                }

                target.SegmentVolume = target.SegmentVolume.Margin(-5 - Radius);

                CreatePrv(latticeParams.OarStructures);
                var prv = structureSet.Structures.FirstOrDefault(x => x.Id == "zzz_SFRT_PRV");

                target.SegmentVolume = target.SegmentVolume.Sub(prv);


                Structure eval_ptv = null;
                CreateStructure(structureSet, "zzz_EvalPTV", false, false);
                eval_ptv = structureSet.Structures.FirstOrDefault(x => x.Id == "zzz_EvalPTV");
                AddContoursToMain(structureSet, ref eval_ptv, ref ptv_low);
                // eval_ptv.SegmentVolume = eval_ptv.SegmentVolume.Sub(prv);

                foreach (var oar in latticeParams.OarStructures)
                {
                    progress?.Report($"\nAccounting for OAR structure: {oar}");
                    var oar_contour = structureSet.Structures.FirstOrDefault(x => x.Id == oar);
                    // Check if oar_contour has contours
                    if (!oar_contour.HasSegment)
                    {
                        continue;
                    }

                    if (oar_contour.IsHighResolution)
                    {
                        CreateStructure(structureSet, "zzz_temp_oar", false, false, "Control");
                        var temp_oar = structureSet.Structures.FirstOrDefault(x => x.Id == "zzz_temp_oar");
                        AddContoursToMain(structureSet, ref temp_oar, ref oar_contour);
                        eval_ptv.SegmentVolume = eval_ptv.SegmentVolume.Sub(temp_oar);
                        structureSet.RemoveStructure(temp_oar);
                    }
                    else
                    {
                        eval_ptv.SegmentVolume = eval_ptv.SegmentVolume.Sub(oar_contour);
                    }
                }

                if (HighRes)
                {
                    eval_ptv.ConvertToHighResolution();
                }


                // Generate a regular grid accross the dummy bounding box 
                var bounds = target.MeshGeometry.Bounds;

                // If alignGrid calculate z to snap to
                double z0 = bounds.Z;
                double zf = bounds.Z + bounds.SizeZ;
                if (alignGrid)
                {
                    // Snap z to nearest z slice
                    // where z slices = img.origin.z + (c * zres)
                    // x, y, z --> dropdown all equal
                    // z0 --> rounded to nearest grid slice
                    var zSlices = new List<double>();
                    var plane_idx = (bounds.Z - structureSet.Image.Origin.z) / structureSet.Image.ZRes;
                    int plane_int = (int)Math.Round(plane_idx);

                    z0 = structureSet.Image.Origin.z + (plane_int * structureSet.Image.ZRes);
                }
                sw.Stop();
                progress?.Report($"\nTime to prepare structures: {sw.ElapsedMilliseconds} ms");
                sw.Reset();
                sw.Start();
                progress?.Report("\nGenerating template lattice grid ...");

                // Get points that are not in the image
                List<VVector> grid = null;
                List<VVector> cold_grid = null;

                // We give extra padding of the hex grid to make sure that we account for spheres when we shift the grid
                var xmin = bounds.X - LatticeSpacing * 1.5;
                var ymin = bounds.Y - LatticeSpacing * 1.5;
                var zmin = bounds.Z - LatticeSpacing * 1.5;
                var xsize = bounds.SizeX + LatticeSpacing * 3.1;
                var ysize = bounds.SizeY + LatticeSpacing * 3.1;
                var zsize = bounds.SizeZ + LatticeSpacing * 3.1;
                // grid = BuildHexGrid(bounds.X + XShift, bounds.SizeX, bounds.Y + YShift, bounds.SizeY, z0, bounds.SizeZ);
                grid = BuildHexGrid(xmin, xsize, ymin, ysize, zmin, zsize);
                structMain = CreateStructure(structureSet, "PTV_Peak", false, HighRes);
                // cold_grid = BuildHexGrid(bounds.X + XShift - LatticeSpacing / 2, bounds.SizeX + LatticeSpacing / 2, bounds.Y + YShift, bounds.SizeY, z0, bounds.SizeZ);

                // put even more padding on the cold grid
                xmin -= LatticeSpacing * 1;
                ymin -= LatticeSpacing * 1;
                zmin -= LatticeSpacing * 1;
                xsize += LatticeSpacing * 2.1;
                ysize += LatticeSpacing * 2.1;
                zsize += LatticeSpacing * 2.1;
                cold_grid = BuildHexGrid(xmin - LatticeSpacing / 2, xsize + LatticeSpacing / 2, ymin, ysize, zmin, zsize);
                structMain_cold = CreateStructure(structureSet, "PTV_Valley", false, HighRes, "Control");

                int max_num_shifts = Convert.ToInt32(Math.Round(LatticeSpacing * 0.5 * LatticeSpacing * LatticeSpacing / resolution));

                int optimal_idx = 0;
                double optimal_avg_dist = 999999999999999;

                var target_centroid = target.CenterPoint;

                sw.Stop();
                progress?.Report($"\nTime to generate grid: {sw.ElapsedMilliseconds} ms");
                sw.Reset();

                sw.Start();
                progress?.Report("\nGenerating target mask ...");

                bool[,,] target_mask = GenerateTargetMask(target);

                // List<VVector> full_search_grid = new List<VVector>();

                sw.Stop();
                progress?.Report($"\nTime to generate target mask: {sw.ElapsedMilliseconds} ms");
                // MessageBox.Show("Ready to search for optimal shift, Press OK to continue.");
                sw.Reset();
                sw.Start();
                progress?.Report("\nSearching for optimal shift ...");

                List<int> max_idx_list = SearchGrid(grid, target_mask, target);

                sw.Stop();
                progress?.Report($"\nTime to find optimal shift: {sw.ElapsedMilliseconds} ms");
                // MessageBox.Show("Optimal shift found, press OK to continue.");
                sw.Reset();
                sw.Start();

                // Find the optimal shift among the max_idx_list
                foreach (var idx in max_idx_list)
                {
                    var shift = index_to_shift(idx);
                    var x_shift = shift.x;
                    var y_shift = shift.y;
                    var z_shift = shift.z;

                    var grid_shifted = grid.Select(x => new VVector(x.x + x_shift, x.y + y_shift, x.z + z_shift)).ToList();

                    grid_shifted = CheckGridInContour(grid_shifted, target);

                    double avg_dist = 0;
                    foreach (var pt in grid_shifted)
                    {
                        avg_dist += VVector.Distance(pt, target_centroid);
                    }
                    avg_dist /= grid_shifted.Count;

                    if (avg_dist == 0)
                    {
                        throw new Exception("No valid points in target");
                    }
                    if (avg_dist < optimal_avg_dist)
                    {
                        optimal_avg_dist = avg_dist;
                        optimal_idx = idx;
                    }
                }

                var opt_shift = index_to_shift(optimal_idx);

                grid = grid.Select(x => new VVector(x.x + opt_shift.x, x.y + opt_shift.y, x.z + opt_shift.z)).ToList();
                cold_grid = cold_grid.Select(x => new VVector(x.x + opt_shift.x, x.y + opt_shift.y, x.z + opt_shift.z)).ToList();

                // Copy grid to grid_template
                var grid_template = grid.ToList();

                // Create new structure
                // Check later
                CreateStructure(structureSet, "zzz_extra", false, HighRes, "Control");
                var structTemplate = structureSet.Structures.FirstOrDefault(x => x.Id == "zzz_extra");


                // Check if grid points are within target
                grid = CheckGridInContour(grid, target);
                cold_grid = CheckGridInContour(cold_grid, eval_ptv);


                cold_grid = CheckColdGrid(cold_grid, grid);
                progress?.Report("\n Number of hot grid points: " + grid.Count);
                progress?.Report("\n Number of cold grid points: " + cold_grid.Count);

                grid_template = CheckTemplateGrid(grid_template, grid);
                grid_template = CheckGridInContour(grid_template, ptv_low);

                sw.Stop();
                progress?.Report($"\nTime to finalize grid: {sw.ElapsedMilliseconds} ms");
                sw.Reset();
                sw.Start();
                progress?.Report($"\nCreating hot spheres ...");

                // Create all individual spheres
                Dictionary<int, int> SliceIdx = new Dictionary<int, int>();
                int idx_tracker = 0;
                grid.Reverse();
                foreach (VVector ctr in grid)
                {
                    int z_slice = (int)Math.Round((ctr.z - structureSet.Image.Origin.z) / structureSet.Image.ZRes);
                    if (SliceIdx.ContainsKey(z_slice))
                    {
                        string structure_name = "TS_Peak_" + SliceIdx[z_slice].ToString();
                        // If the slice already exists, just add to it
                        var ts_peak = structureSet.Structures.FirstOrDefault(x => x.Id == structure_name);
                        BuildSphere(ts_peak, ctr, Radius);
                    }
                    else
                    {
                        SliceIdx.Add(z_slice, idx_tracker);
                        string structure_name = "TS_Peak_" + SliceIdx[z_slice].ToString();
                        idx_tracker++;
                        var ts_peak = CreateStructure(structureSet, structure_name, false, HighRes, "PTV");
                        BuildSphere(ts_peak, ctr, Radius);
                    }
                }
                foreach (int idx in SliceIdx.Keys)
                {
                    string structure_name = "TS_Peak_" + SliceIdx[idx].ToString();
                    var ts_peak = structureSet.Structures.FirstOrDefault(x => x.Id == structure_name);
                    structMain.SegmentVolume = structMain.SegmentVolume.Or(ts_peak.SegmentVolume);
                }

                SliceIdx.Clear();
                idx_tracker = 0;
                cold_grid.Reverse();
                progress?.Report($"\nCreating cold spheres ...");
                foreach (VVector ctr in cold_grid)
                {
                    int z_slice = (int)Math.Round((ctr.z - structureSet.Image.Origin.z) / structureSet.Image.ZRes);
                    if (SliceIdx.ContainsKey(z_slice))
                    {
                        string structure_name = "TS_Valley_" + SliceIdx[z_slice].ToString();
                        // If the slice already exists, just add to it
                        var ts_valley = structureSet.Structures.FirstOrDefault(x => x.Id == structure_name);
                        BuildSphere(ts_valley, ctr, Radius);
                        structMain_cold.SegmentVolume = structMain_cold.SegmentVolume.Or(ts_valley.SegmentVolume);
                    }
                    else
                    {
                        SliceIdx.Add(z_slice, idx_tracker);
                        string structure_name = "TS_Valley_" + SliceIdx[z_slice].ToString();
                        idx_tracker++;
                        var ts_valley = CreateStructure(structureSet, structure_name, false, HighRes, "Control");
                        BuildSphere(ts_valley, ctr, Radius);
                        structMain_cold.SegmentVolume = structMain_cold.SegmentVolume.Or(ts_valley.SegmentVolume);
                    }
                }
                foreach (int idx in SliceIdx.Keys)
                {
                    string structure_name = "TS_Valley_" + SliceIdx[idx].ToString();
                    var ts_valley = structureSet.Structures.FirstOrDefault(x => x.Id == structure_name);
                    ts_valley.SegmentVolume = ts_valley.SegmentVolume.And(eval_ptv);

                    structMain_cold.SegmentVolume = structMain_cold.SegmentVolume.Or(ts_valley.SegmentVolume);
                }

                progress?.Report($"\nCreating extra spheres ...");
                foreach (VVector ctr in grid_template)
                {
                    BuildSphere(structTemplate, ctr, Radius);
                }

                structMain_cold.SegmentVolume = structMain_cold.SegmentVolume.And(eval_ptv);

                // Delete the autogenerated target if it exists
                if (deleteAutoTarget)
                {
                    structureSet.RemoveStructure(target_initial_temp);
                }
                // structureSet.RemoveStructure(eval_ptv);
                // structureSet.RemoveStructure(target);


                progress?.Report($"\nCreating tuning structures ...");

                // Create PTV_Control
                CreateStructure(structureSet, "PTV_Control", false, HighRes, "Control");
                Structure ptv_control = structureSet.Structures.FirstOrDefault(x => x.Id == "PTV_Control");
                AddContoursToMain(structureSet, ref ptv_control, ref ptv_low);
                ptv_control.SegmentVolume = ptv_control.SegmentVolume.Sub(structMain.SegmentVolume.Margin(5));

                // Create TS_Ring
                CreateStructure(structureSet, "TS_Ring2", false, false, "Control");
                Structure ts_ring2 = structureSet.Structures.FirstOrDefault(x => x.Id == "TS_Ring2");
                AddContoursToMain(structureSet, ref ts_ring2, ref ptv_low);
                ts_ring2.SegmentVolume = ts_ring2.SegmentVolume.Margin(30);

                CreateStructure(structureSet, "TS_Ring1", false, false, "Control");
                Structure ts_ring1 = structureSet.Structures.FirstOrDefault(x => x.Id == "TS_Ring1");
                AddContoursToMain(structureSet, ref ts_ring1, ref ptv_low);
                ts_ring1.SegmentVolume = ts_ring1.SegmentVolume.Margin(15);

                if (ptv_low.IsHighResolution)
                {
                    CreateStructure(_ss, "zzz_temp_ptv_low", false, false, "Control");
                    var temp_ptv_low = _ss.Structures.FirstOrDefault(x => x.Id == "zzz_temp_ptv_low");
                    AddContoursToMain(_ss, ref temp_ptv_low, ref ptv_low);

                    ts_ring2.SegmentVolume = ts_ring2.SegmentVolume.Sub(temp_ptv_low.SegmentVolume.Margin(15));
                    ts_ring1.SegmentVolume = ts_ring1.SegmentVolume.Sub(temp_ptv_low.SegmentVolume.Margin(5));

                    structureSet.RemoveStructure(temp_ptv_low);
                }
                else
                {
                    ts_ring2.SegmentVolume = ts_ring2.SegmentVolume.Sub(ptv_low.SegmentVolume.Margin(15));
                    ts_ring1.SegmentVolume = ts_ring1.SegmentVolume.Sub(ptv_low.SegmentVolume.Margin(5));
                }
                    
                var body = structureSet.Structures.FirstOrDefault(x => x.Id == BodyId);
                ts_ring2.SegmentVolume = ts_ring2.SegmentVolume.And(body);
                ts_ring1.SegmentVolume = ts_ring1.SegmentVolume.And(body);

                CreateStructure(structureSet, "TS_Peak_Ring1", false, HighRes, "Control");
                Structure ts_peak_ring1 = structureSet.Structures.FirstOrDefault(x => x.Id == "TS_Peak_Ring1");
                AddContoursToMain(structureSet, ref ts_peak_ring1, ref structMain);
                ts_peak_ring1.SegmentVolume = ts_peak_ring1.SegmentVolume.Margin(5);
                ts_peak_ring1.SegmentVolume = ts_peak_ring1.SegmentVolume.Sub(structMain.SegmentVolume.Margin(0));


                CreateStructure(structureSet, "TS_Peak_Ring2", false, HighRes, "Control");
                Structure ts_peak_ring2 = structureSet.Structures.FirstOrDefault(x => x.Id == "TS_Peak_Ring2");
                AddContoursToMain(structureSet, ref ts_peak_ring2, ref structMain);
                ts_peak_ring2.SegmentVolume = ts_peak_ring2.SegmentVolume.Margin(10);
                ts_peak_ring2.SegmentVolume = ts_peak_ring2.SegmentVolume.Sub(structMain.SegmentVolume.Margin(5));

                CreateStructure(structureSet, "TS_Peak_Ring3", false, HighRes, "Control");
                Structure ts_peak_ring3 = structureSet.Structures.FirstOrDefault(x => x.Id == "TS_Peak_Ring3");
                AddContoursToMain(structureSet, ref ts_peak_ring3, ref structMain);
                ts_peak_ring3.SegmentVolume = ts_peak_ring3.SegmentVolume.Margin(15);
                ts_peak_ring3.SegmentVolume = ts_peak_ring3.SegmentVolume.Sub(structMain.SegmentVolume.Margin(10));

                progress?.Report("\nLattice structure creation complete!");
                sw.Stop();
                progress?.Report($"\nTime to create spheres: {sw.ElapsedMilliseconds} ms");

            });
        }

        private VVector[] CreateContour(VVector center, double radius, int nOfPoints)
        {
            VVector[] contour = new VVector[nOfPoints + 1];
            double angleIncrement = Math.PI * 2.0 / Convert.ToDouble(nOfPoints);
            for (int i = 0; i < nOfPoints; ++i)
            {
                double angle = Convert.ToDouble(i) * angleIncrement;
                double xDelta = radius * Math.Cos(angle);
                double yDelta = radius * Math.Sin(angle);
                VVector delta = new VVector(xDelta, yDelta, 0.0);
                contour[i] = center + delta;
            }
            contour[nOfPoints] = contour[0];

            return contour;
        }


        public async Task SetupBeams(LatticeParameters latticeParams)
        {
            await _ew.AsyncRunPlanContext((pat, ps) =>
            {
                var plan = (ExternalPlanSetup)ps;
                StructureSet structureSet = ps.StructureSet;

                // Clean up all beams
                var all_beam = plan.Beams.ToList();

                foreach (var beam in all_beam)
                {
                    plan.RemoveBeam(beam);
                }

                var ptv_low = structureSet.Structures.FirstOrDefault(x => x.Id == latticeParams.PTVLowStructure);
                VVector isocenter = ptv_low.CenterPoint;
                // round isocenter to 1 decimal place
                isocenter = new VVector(Math.Round(isocenter.x, 0), Math.Round(isocenter.y, 0), Math.Round(isocenter.z, 0));

                string machineId = latticeParams.MachineId;
                string energy = latticeParams.Energy;
                int doseRate = 600;
                string primaryFluenceModeId = "";
                if (energy.Contains("FFF"))
                {
                    primaryFluenceModeId = "FFF";
                    energy = energy.Replace("-FFF", "");
                    if (energy.Contains("6X"))
                    {
                        doseRate = 1400;
                    }
                    else if (energy.Contains("10X"))
                    {
                        doseRate = 2400;
                    }
                }
                ExternalBeamMachineParameters beamParams = new ExternalBeamMachineParameters(machineId, energy, doseRate, "SRS ARC", primaryFluenceModeId);
                // make a list of 180 array of zeros
                IEnumerable<double> metersetWeights = Enumerable.Repeat(0.0, 180).ToArray();

                var collimatorAngle = new List<double> { 50.0, 85.0, 125.0 };
                var gantryAngle = new List<double> { 181.0, 179.0, 181.0 };
                var gantryStop = new List<double> { 179.0, 181.0, 179.0 };
                var gantryDirection = VMS.TPS.Common.Model.Types.GantryDirection.Clockwise;
                var couchAngle = new List<double> { 0.0, 0.0, 0.0 };

                if (latticeParams.CouchKick)
                {
                    collimatorAngle = new List<double> { 312.0, 50.0, 57.0 };
                    couchAngle = new List<double> { 350.0, 0.0, 10.0 };
                }

                var jawPositions = new VRect<double>(-100, -100, 100, 100);

                for (int i = 0; i < 3; i++)
                {
                    // beam = plan.AddVMATBeam(
                    // beamParams,
                    // metersetWeights,
                    // Math.Round(collimatorAngle[i], 1),
                    // gantryAngle[i],
                    // gantryStop[i],
                    // gantryDirection,
                    // couchAngle[i],
                    // isocenter);
                    // CloseModalWindows();
                    var beam = plan.AddMLCArcBeam(
                        beamParams,
                        null,
                        jawPositions,
                        Math.Round(collimatorAngle[i], 1),
                        gantryAngle[i],
                        gantryStop[i],
                        gantryDirection,
                        couchAngle[i],
                        isocenter);

                    if (gantryDirection == VMS.TPS.Common.Model.Types.GantryDirection.Clockwise)
                    {

                        beam.Id = (i + 1).ToString() + " CW";
                    }
                    else
                    {
                        beam.Id = (i + 1).ToString() + " CCW";
                    }


                    gantryDirection = gantryDirection == VMS.TPS.Common.Model.Types.GantryDirection.Clockwise
                        ? VMS.TPS.Common.Model.Types.GantryDirection.CounterClockwise
                        : VMS.TPS.Common.Model.Types.GantryDirection.Clockwise;

                    var fitmargin = new VMS.TPS.Common.Model.Types.FitToStructureMargins(5.0);
                    var meetingpoint = VMS.TPS.Common.Model.Types.OpenLeavesMeetingPoint.OpenLeavesMeetingPoint_Middle;
                    var closedMeetingPoint = VMS.TPS.Common.Model.Types.ClosedLeavesMeetingPoint.ClosedLeavesMeetingPoint_Center;
                    var jawfitting = VMS.TPS.Common.Model.Types.JawFitting.FitToStructure;
                    beam.FitMLCToStructure(fitmargin, ptv_low, false, jawfitting, meetingpoint, closedMeetingPoint);
                }
            });
        }

        public async Task SetupOptimizer(LatticeParameters latticeParams)
        {
            await _ew.AsyncRunPlanContext((pat, ps) =>
            {
                var plan = (ExternalPlanSetup)ps;
                StructureSet structureSet = ps.StructureSet;

                string axb_model = "AcurosXB_1811";
                plan.SetCalculationOption(axb_model, "CalculationGridSizeInCM", "0.25");
                plan.SetCalculationOption(axb_model, "UseGPU", "No");

                Optimizer = plan.OptimizationSetup;

                // Clean up the objectives:
                var all_objectives = Optimizer.Objectives;
                foreach (var objective in all_objectives)
                {
                    Optimizer.RemoveObjective(objective);
                }

                Optimizer.UseJawTracking = true;
                Optimizer.AddAutomaticSbrtNormalTissueObjective(80.0);

                var ptv_low_name = latticeParams.PTVLowStructure;
                var ptv_low = structureSet.Structures.FirstOrDefault(x => x.Id == ptv_low_name);
                var doseObjective = new DoseValue(2000.0, "cGy");
                if (ptv_low != null)
                {
                    Optimizer.AddPointObjective(ptv_low, OptimizationObjectiveOperator.Lower, doseObjective, 100.0, 100.0);
                }

                List<Structure> peak_structures = new List<Structure>();
                List<Structure> valley_structures = new List<Structure>();

                foreach (var struct_i in structureSet.Structures)
                {
                    if (struct_i.Id.Contains("TS_Peak_") && !struct_i.Id.Contains("Ring"))
                    {
                        peak_structures.Add(struct_i);
                    }
                    else if (struct_i.Id.Contains("TS_Valley_"))
                    {
                        valley_structures.Add(struct_i);
                    }
                }

                foreach (var peak_structure in peak_structures)
                {
                    AddPeakObjective(peak_structure, Optimizer);
                }
                foreach (var valley_structure in valley_structures)
                {
                    AddValleyObjective(valley_structure, Optimizer);
                }

                // Add the TS_Peak_Ring objectives
                doseObjective = new DoseValue(6670.0, "cGy");
                var ts_peak_ring1 = structureSet.Structures.FirstOrDefault(x => x.Id == "TS_Peak_Ring1");
                Optimizer.AddPointObjective(ts_peak_ring1, OptimizationObjectiveOperator.Upper, doseObjective, 0.0, 30);

                doseObjective = new DoseValue(6000.0, "cGy");
                var ts_peak_ring2 = structureSet.Structures.FirstOrDefault(x => x.Id == "TS_Peak_Ring2");
                Optimizer.AddPointObjective(ts_peak_ring2, OptimizationObjectiveOperator.Upper, doseObjective, 0.0, 30);

                doseObjective = new DoseValue(5000.0, "cGy");
                var ts_peak_ring3 = structureSet.Structures.FirstOrDefault(x => x.Id == "TS_Peak_Ring3");
                Optimizer.AddPointObjective(ts_peak_ring3, OptimizationObjectiveOperator.Upper, doseObjective, 0.0, 30);

                // Add TS_Ring objectives
                doseObjective = new DoseValue(2000.0, "cGy");
                var ts_ring1 = structureSet.Structures.FirstOrDefault(x => x.Id == "TS_Ring1");
                Optimizer.AddPointObjective(ts_ring1, OptimizationObjectiveOperator.Upper, doseObjective, 0.0, 30);

                doseObjective = new DoseValue(1500.0, "cGy");
                var ts_ring2 = structureSet.Structures.FirstOrDefault(x => x.Id == "TS_Ring2");
                Optimizer.AddPointObjective(ts_ring2, OptimizationObjectiveOperator.Upper, doseObjective, 0.0, 30);

                // plan.OptimizeVMAT();
            });


            // I would do one round of optimization, then find all organs-at-risk that are within PTV bounding box +/- 5 cm in sup-inf
            // Apply a mean dose reduction
            // Apply the constraints that we have in CTP - 20% 
            // 
        }

        private void AddPeakObjective(Structure peak_structure, OptimizationSetup optimizer)
        {
            if (peak_structure == null)
            {
                return;
            }
            var upperDose = new DoseValue(7000.0, "cGy");
            optimizer.AddPointObjective(peak_structure, OptimizationObjectiveOperator.Upper, upperDose, 0.0, 100.0);
            var lowerDose = new DoseValue(6670.0, "cGy");
            optimizer.AddPointObjective(peak_structure, OptimizationObjectiveOperator.Lower, lowerDose, 100.0, 100.0);
            var lowerDose2 = new DoseValue(6670.0, "cGy");
            optimizer.AddPointObjective(peak_structure, OptimizationObjectiveOperator.Lower, lowerDose2, 95.0, 130.0);
        }

        private void AddValleyObjective(Structure valley_structure, OptimizationSetup optimizer)
        {
            if (valley_structure == null)
            {
                return;
            }
            var upperDose = new DoseValue(2400.0, "cGy");
            optimizer.AddPointObjective(valley_structure, OptimizationObjectiveOperator.Upper, upperDose, 0.0, 100.0);
            var lowerDose = new DoseValue(1950.0, "cGy");
            optimizer.AddPointObjective(valley_structure, OptimizationObjectiveOperator.Lower, lowerDose, 100.0, 100.0);
            var eudDose = new DoseValue(2100.0, "cGy");
            optimizer.AddEUDObjective(valley_structure, OptimizationObjectiveOperator.Exact, eudDose, -5.0, 70.0);
        }

        // --- begin added: optimization parameter loader and storage ---
        private List<OptimizationParameter> _optimizationParameters = new List<OptimizationParameter>();

        // keep track of structures that received template objectives
        private List<string> OAR_list = new List<string>();

        private class OptimizationParameter
        {
            public string StructureType { get; set; }
            public List<string> Labels { get; set; } = new List<string>();
            public string ObjectiveType { get; set; } // e.g. "dvh" or "mean"
            public string ObjectiveOperator { get; set; } // e.g. "Upper", "Lower", "Exact"
            public string ObjectiveParameter { get; set; } // e.g. "Mean", "V2000", "D20%"
            public string ObjectiveValue { get; set; } // value in cGy or cc
        }

        /// <summary>
        /// Load optimization parameters from CSV. Call this before SetupOptimizer().
        /// CSV header: StructureType,Labels,ObjectiveType,ObjectiveOperator,ObjectiveParameter,ObjectiveValue
        /// Labels are semicolon separated substrings used to match structure.Id
        /// </summary>
        public void LoadOptimizationParameters(string csvPath)
        {
            _optimizationParameters.Clear();
            if (!File.Exists(csvPath)) return;

            var lines = File.ReadAllLines(csvPath);
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i].Trim();
                if (string.IsNullOrEmpty(line)) continue;
                if (i == 0 && line.ToLower().StartsWith("structuretype")) continue;

                // naive CSV split (no quoted fields). Good for simple templates.
                var parts = line.Split(new[] { ',' }, StringSplitOptions.None);
                if (parts.Length < 6) continue;

                var param = new OptimizationParameter
                {
                    StructureType = parts[0].Trim(),
                    Labels = parts[1].Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList(),
                    ObjectiveType = parts[2].Trim().ToLower(),
                    ObjectiveOperator = parts[3].Trim(),
                    ObjectiveParameter = parts[4].Trim(),
                    ObjectiveValue = parts[5].Trim()
                };

                _optimizationParameters.Add(param);
            }
        }

        /// <summary>
        /// Load the bundled CSV from the application's output folder (SFRT_configs\objective_template.csv).
        /// Call this once at startup (before SetupOptimizer).
        /// </summary>
        public void LoadDefaultOptimizationParameters()
        {
            // runtime folder (bin\Debug|Release\netX\)
            var AssemblyLocation = Assembly.GetExecutingAssembly().Location;
            var baseDir = Path.GetDirectoryName(AssemblyLocation);
            // var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var cfgPath = Path.Combine(baseDir, "SFRT_configs", "objective_template.csv");

            if (File.Exists(cfgPath))
            {
                LoadOptimizationParameters(cfgPath);
            }
            else
            {
                MessageBox.Show("Default optimization parameter template not found:\n" + cfgPath);
                // optional: silently ignore or throw/log; kept silent to avoid ESAPI context issues
            }
        }

        private OptimizationObjectiveOperator ParseOperatorOrDefault(string opStr, OptimizationObjectiveOperator defaultOp)
        {
            if (string.IsNullOrWhiteSpace(opStr)) return defaultOp;
            if (Enum.TryParse<OptimizationObjectiveOperator>(opStr, true, out var parsed))
            {
                return parsed;
            }
            return defaultOp;
        }

        // helper to apply loaded parameters to structures
        private void ApplyLoadedObjectives(OptimizationSetup optimizer)
        {
            if (_optimizationParameters == null || _optimizationParameters.Count == 0) return;

            // clear previous tracking list
            OAR_list.Clear();

            foreach (var opt in _optimizationParameters)
            {
                foreach (var label in opt.Labels)
                {
                    var matches = _ss.Structures
                        .Where(s => s.Id.Equals(label, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    foreach (var s in matches)
                    {
                        try
                        {
                            bool addedForStructure = false;

                            // decide operator default: Upper
                            var defaultOp = OptimizationObjectiveOperator.Upper;
                            var opToUse = ParseOperatorOrDefault(opt.ObjectiveOperator, defaultOp);
                            double weight = 30.0;


                            if (opt.ObjectiveType == "dvh")
                            {
                                // Parse if it's a volume or dose based objective
                                if (opt.ObjectiveParameter.StartsWith("d", StringComparison.OrdinalIgnoreCase))
                                {
                                    double doseDouble = 0;
                                    double.TryParse(opt.ObjectiveValue, out doseDouble);
                                    var dv = new DoseValue(doseDouble, "cGy");
                                    if (opt.ObjectiveParameter.EndsWith("cc"))
                                    {
                                        // e.g. D20cc
                                        double volCc = 0;
                                        var volStr = opt.ObjectiveParameter.Substring(1).Replace("cc", "");
                                        double.TryParse(volStr, out volCc);
                                        double volPercent = volCc / s.Volume * 100.0;
                                        optimizer.AddPointObjective(s, opToUse, dv, volPercent, weight);
                                        addedForStructure = true;
                                    }
                                    else
                                    {
                                        // e.g. D20%
                                        double volPercent = 0;
                                        var volStr = opt.ObjectiveParameter.Substring(1).Replace("%", "");
                                        double.TryParse(volStr, out volPercent);
                                        optimizer.AddPointObjective(s, opToUse, dv, volPercent, weight);
                                        addedForStructure = true;
                                    }
                                }
                                else if (opt.ObjectiveParameter.StartsWith("v", StringComparison.OrdinalIgnoreCase))
                                {
                                    // Check if it's volume spared or volume treated
                                    if (opt.ObjectiveParameter.StartsWith("vs", StringComparison.OrdinalIgnoreCase))
                                    {
                                        string doseString = opt.ObjectiveParameter.Substring(2);
                                        double dosecGy = 0;
                                        double.TryParse(doseString, out dosecGy);
                                        var dv = new DoseValue(dosecGy, "cGy");

                                        if (opt.ObjectiveValue.EndsWith("cc"))
                                        {
                                            string volString = opt.ObjectiveValue.Substring(2).Replace("cc", "");
                                            double volCc = 0;
                                            double.TryParse(volString, out volCc);

                                            double structureVolume = s.Volume;
                                            double volSparedPercent = (1.0 - volCc / structureVolume) * 100.0;

                                            optimizer.AddPointObjective(s, opToUse, dv, volSparedPercent, weight);
                                            addedForStructure = true;
                                        }
                                        else
                                        {
                                            string volString = opt.ObjectiveValue.Substring(2).Replace("%", "");
                                            double volPercent = 0;
                                            double.TryParse(volString, out volPercent);
                                            double volSparedPercent = 100.0 - volPercent;
                                            optimizer.AddPointObjective(s, opToUse, dv, volSparedPercent, weight);
                                            addedForStructure = true;
                                        }
                                    }
                                    else
                                    {
                                        string doseString = opt.ObjectiveParameter.Substring(1);
                                        double dosecGy = 0;
                                        double.TryParse(doseString, out dosecGy);
                                        var dv = new DoseValue(dosecGy, "cGy");

                                        if (opt.ObjectiveValue.EndsWith("cc"))
                                        {
                                            var volStr = opt.ObjectiveValue.Replace("cc", "");
                                            double volCc = 0;
                                            double.TryParse(volStr, out volCc);
                                            double volPercent = volCc / s.Volume * 100.0;
                                            optimizer.AddPointObjective(s, opToUse, dv, volPercent, weight);
                                            addedForStructure = true;
                                        }
                                        else
                                        {
                                            var volStr = opt.ObjectiveValue.Replace("%", "");
                                            double volPercent = 0;
                                            double.TryParse(volStr, out volPercent);
                                            optimizer.AddPointObjective(s, opToUse, dv, volPercent, weight);
                                            addedForStructure = true;
                                        }

                                    }
                                }
                                else
                                {
                                    // raise exception
                                    throw new Exception("Unknown ObjectiveParameter format: " + opt.ObjectiveParameter);
                                }

                                // Use AddPointObjective with provided operator (Upper/Lower)
                                // optimizer.AddPointObjective(s, opToUse, dv, 0.0, weight);
                            }
                            else if (opt.ObjectiveType == "mean")
                            {
                                double doseDouble = 0;
                                double.TryParse(opt.ObjectiveValue, out doseDouble);
                                var dv = new DoseValue(doseDouble, "cGy");
                                // map mean to EUD-style objective; operator typically Exact but use provided if sensible
                                optimizer.AddEUDObjective(s, opToUse, dv, -1.0, weight);
                                addedForStructure = true;
                            }
                            else
                            {
                                // unknown objective type
                                throw new Exception("Unknown ObjectiveType: " + opt.ObjectiveType);
                            }

                            // track that we added at least one objective for this structure
                            if (addedForStructure && !OAR_list.Contains(s.Id, StringComparer.OrdinalIgnoreCase))
                            {
                                OAR_list.Add(s.Id);
                            }
                        }
                        catch (Exception)
                        {
                            // swallow to avoid crashing ESAPI context; optionally log error
                        }
                    }
                    if (matches.Count > 0)
                    {
                        // structure matched, no need to check other labels for this parameter
                        break;
                    }
                }
            }
            // After applying template objectives, reduce dose constraints by 5% for objectives associated with tracked OARs.
            // copy list as optimizer.Objectives may be a live collection
            var objectivesCopy = optimizer.Objectives.ToList();
            foreach (var obj in objectivesCopy)
            {
                try
                {
                    // get the structure id for this objective (best-effort via reflection)
                    string structId = null;
                    var structProp = obj.GetType().GetProperty("Structure");
                    if (structProp != null)
                    {
                        var structObj = structProp.GetValue(obj);
                        if (structObj != null)
                        {
                            var idProp = structObj.GetType().GetProperty("Id");
                            structId = idProp?.GetValue(structObj)?.ToString();
                        }
                    }
                    if (string.IsNullOrEmpty(structId))
                    {
                        var idProp2 = obj.GetType().GetProperty("StructureId") ?? obj.GetType().GetProperty("StructureName");
                        structId = idProp2?.GetValue(obj)?.ToString();
                    }

                    if (string.IsNullOrEmpty(structId)) continue;
                    if (!OAR_list.Contains(structId, StringComparer.OrdinalIgnoreCase)) continue;

                    // Gather a reference to the Structure object
                    Structure structRef = null;
                    if (structProp != null)
                    {
                        structRef = structProp.GetValue(obj) as Structure;
                    }
                    if (structRef == null)
                    {
                        structRef = _ss.Structures.FirstOrDefault(s => s.Id.Equals(structId, StringComparison.OrdinalIgnoreCase));
                    }
                    if (structRef == null) continue;

                    // Get operator (best-effort)
                    OptimizationObjectiveOperator op = OptimizationObjectiveOperator.Upper;
                    var opProp = obj.GetType().GetProperty("Operator") ?? obj.GetType().GetProperty("ObjectiveOperator");
                    if (opProp != null)
                    {
                        try
                        {
                            var opObj = opProp.GetValue(obj);
                            if (opObj != null) Enum.TryParse(opObj.ToString(), true, out op);
                        }
                        catch { }
                    }

                    // get volume and weight (fallbacks)
                    double volume = 0.0;
                    double weight = 30.0;
                    var volProp = obj.GetType().GetProperty("Volume") ?? obj.GetType().GetProperty("VolumeFraction");
                    if (volProp != null) double.TryParse(volProp.GetValue(obj)?.ToString(), out volume);
                    var wtProp = obj.GetType().GetProperty("Weight");
                    if (wtProp != null) double.TryParse(wtProp.GetValue(obj)?.ToString(), out weight);

                    // read existing dose value (best-effort)
                    double existingDoseValue = 0.0;
                    string existingDoseUnit = "cGy";
                    var doseProp = obj.GetType().GetProperty("Dose") ?? obj.GetType().GetProperty("DoseValue");
                    if (doseProp != null && doseProp.CanRead)
                    {
                        var doseObj = doseProp.GetValue(obj);
                        if (doseObj != null)
                        {
                            var dValProp = doseObj.GetType().GetProperty("Dose");
                            var dUnitProp = doseObj.GetType().GetProperty("Unit");
                            if (dValProp != null)
                            {
                                double.TryParse(dValProp.GetValue(doseObj)?.ToString(), out existingDoseValue);
                            }
                            if (dUnitProp != null)
                            {
                                existingDoseUnit = dUnitProp.GetValue(doseObj)?.ToString() ?? existingDoseUnit;
                            }
                        }
                    }

                    // remove and re-add with 5% lower dose if we have a dose value
                    if (existingDoseValue > 0)
                    {
                        var newDoseVal = new DoseValue(existingDoseValue * 0.95, existingDoseUnit);
                        try
                        {
                            optimizer.RemoveObjective(obj);
                            optimizer.AddPointObjective(structRef, op, newDoseVal, volume, weight);
                        }
                        catch
                        {
                            // swallow to avoid ESAPI context errors
                        }
                    }
                }
                catch
                {
                    // ignore individual objective failures
                }
            }
        }

        public async Task OptimizeLattice(IProgress<string> progress = null, LatticeParameters latticeParams = null)
        {
            if (latticeParams == null)
            {
                throw new ArgumentNullException(nameof(latticeParams), "Lattice parameters must be provided.");
            }

            await SetupOptimizer(latticeParams);
            await _ew.AsyncRunPlanContext((pat, ps) =>
            {
                // This is required to update the structure set reference after changes
                _ss = ps.StructureSet;
                progress?.Report("\nLattice optimization setup complete.");
                ApplyLoadedObjectives(Optimizer);
                progress?.Report("\nApplied template objectives.");
                // var plan = (ExternalPlanSetup)ps;
                // plan.OptimizeVMAT();
                // progress?.Report("\nOptimization complete. Test");
            });
        }
    }
}