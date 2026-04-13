# SFRT Automated Planning Scripts for RayStation

These scripts automate Spatially Fractionated Radiation Therapy (SFRT) lattice planning in RayStation. Updated for **RayStation v14**.

The three scripts are intended to be run in order:

1. `SFRT_Target_Contours.py` - Generate lattice target contours
2. `SFRT_Plan_Creation.py` - Create the treatment plan
3. `SFRT_Dose_Optimization.py` - Optimize the dose

All CSV configuration files must be placed in the same directory as the scripts.

---

## SFRT_Target_Contours

Generates the lattice sphere pattern and target contour structures from a GTV and PTV.

### Required inputs

- A patient with at least one CT examination
- GTV structure
- PTV structure (e.g. PTV_LD)
- External (body) contour
- Optional: OAR structures with PRV margins, specified via the GUI

### Outputs

- **VTV_HD**: hot sphere lattice intersected with the GTV.
- **VTV_LD**: cold sphere lattice intersected with the PTV
- **ICRU_X**: POI placed at the center of the hot sphere closest to the GTV centroid

### Notes

- The GUI allows selection of the examination, GTV, PTV, external contour, and lattice configuration (default, 1.5 cm, or 1.0 cm spheres).
- OAR avoidance structures can be added with custom margins. The lattice will avoid these regions.
- If fewer than 6 spheres fit (default threshold), the script automatically retries with smaller 1.0 cm spheres and tighter spacing.

---

## SFRT_Plan_Creation

### Required inputs

- Isocenter point must be positioned and named ahead of time. The script will ask you to select the isocenter and will not modify it.
- **ICRU_X** must exist. This point is created by SFRT_Target_Contours.
- **PTV_LD** must be present to create Eval_PTV_LD. If Eval_PTV_LD already exists, PTV_LD is not needed.
- **VTV_HD**, created by SFRT_Target_Contours.
- `SFRT_ClinicalGoals_[15mm|10mm]_[20Gy|10Gy].csv` and `SFRT_ClinicalGoals_OAR.csv` files in the same directory as the script.

### Outputs

- Plan and beam sets with the correct MOSAIQ letter suffix. Optimization settings: 80 max iterations, 30 before conversion.
- Clinical goals added to the plan. Only goals matching an existing structure are added.
- **ICRU\_??** (where ?? is the MOSAIQ letter): a copy of ICRU_X.
- **Eval_PTV_LD**: PTV_LD cropped from External by 5 mm (if not already present).
- **Skin**: External - 5 mm (if not already present).
- Optimization structures (if not already present):
  - **PTV_Control**
  - **OPT_VTV_HDplus2mm**
  - **x_Ring1**: 0.1 cm - 1.5 cm
  - **x_Ring2**: 1.5 cm - 3 cm
  - **x_Ring3**: 3 cm - 6 cm

### Note: Overwrite contours button

This button deletes all existing optimization structures (PTV_Control, OPT_VTV_HDplus2mm, x_Ring1, x_Ring2, x_Ring3) and recreates them. It does not create a plan — the new plan button must be pressed afterwards.

---

## SFRT_Dose_Optimization

### Required inputs

- All target and optimization structures: **VTV_HD**, **VTV_LD**, **PTV_Control**, **OPT_VTV_HDplus2mm**, **GTV**, **Skin**, and **Eval_PTV_LD**. The script verifies these exist with non-empty contours and displays a message if any are missing.
- Plan must be correctly selected before running the script.
- `SFRT_ObjectiveStep[1,2,3]_[15mm|10mm]_[20Gy|10Gy].csv`, `SFRT_Objective_OAR.csv`, and `SFRT_ClinicalGoals_OAR.csv` files in the same directory as the script.

### Outputs

- Objectives are added according to the CSV spreadsheets. They can be modified prior to optimization by pausing the script with the **Pause script** button.
- Step 1 includes a preliminary optimization to find the max MU/fx limits per beam. The minimum MU/fx per beam is 800 MU.
- Each step optimization runs in 2 sub-runs. The first uses the CSV objectives exactly. After the first run, temporary clinical goals evaluate whether objectives are met. For unmet objectives, weights are increased per the CSV. A 2nd run then executes. The script pauses after this 2nd optimization.
- The dose grid extent along the sup-inf direction (z) is limited to slices containing a relevant structure (excluding Skin) or within 5 cm of PTV_LD.
- Final dose calculation: resets the dose grid to default size, performs final dose calculation, then sets the prescription level to match the isodose level at the ICRU point.
