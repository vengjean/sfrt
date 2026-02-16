SFRT_Plan_Creation
Required inputs
•	Isocenter point needs to be positioned (and named) ahead of time. The script will ask you to select the isocenter and will not modify it in any way.
•	ICRU_X needs to exist. This point is created and positioned by the SFRT_Target_Contours script, which automatically places the POI inside a hot sphere. If the point does not exist, the SFRT_Target_Contours script may not have been run properly.
•	PTVm_2000 (or PTVm_1000 for 10Gy low dose) needs to be present to create Eval_PTVm_2000. If Eval_PTVm_2000 (or Eval_PTVm_1000) is already created, then PTVm_2000 (or 1000) is not needed.
•	PTVm_6670, created by SFRT_Target_Contours. MUST REMOVE SUFFIX
•	SFRT_ClinicalGoals_[15mm or 10mm]_[20Gy or 10Gy].csv files in the Q:\RayStation\Scripting\SFRT folder. This cannot be modified without changing the code.
Outputs
•	Plan and beam sets with the correct MOSAIQ letter suffix. The plan optimization settings are also set: 80 max iterations and 30 before conversion
•	Clinical goals are added to the plan. Only goals in the SFRT_ClinicalGoals spreadsheets that match an existing structure are added.
•	ICRU_??, where ?? is the MOSAIQ letter of the plan. This POI is simply a copy of ICRU_X
•	Eval_PTVm_2000 (or 1000) by cropping it from External by 5 mm, if not already present
•	Skin as External – 5mm, if not already present
•	Optimization structures (Eval_PTVm_Control, x_PTVm_6670Plus2mm, x_Ring1, x_Ring2, x_Ring3), if not already present. 
o	x_Ring1: 0.1 cm  1.5 cm, 
o	x_Ring2: 1.5 cm   3 cm,
o	x_Ring3: 3 cm  6 cm.
Note: [Overwrite contours] button
This button will delete all existing optimization structures (Eval_PTVm_Control, x_PTVm_6670Plus2mm, x_Ring1, x_Ring2, x_Ring3) and recreate them with the script. This button does not create a plan: the new plan button must be pressed afterwards.
SFRT_Dose_Optimization
Required inputs 
•	All target and optimization structures: PTVm_6670, Eval_PTVm_Avoid, Eval_PTVm_Control, x_PTVm_6670Plus2mm, GTVm, Skin and Eval_PTVm_2000 (or 1000) 
The script will verify that all of these structures exist and have non-empty contours. A notification message will appear if any of them are missing. 
These structures are all created from the previous 2 scripts so if any of them are missing, something went wrong in either of the previous scripts.
•	Plan must be correctly selected. I have not found a way to make the script change the plan selection by itself so the plan must be manually selected by the planner.
•	SFRT_ObjectiveStep[1,2 and 3]_[15mm or 10mm]_[20Gy or 10Gy].csv files in the Q:\RayStation\Scripting\SFRT folder. This cannot be modified without changing the code.
Ouputs
•	Objectives are added according to the .csv spreadsheets. They can be modified prior to running the optimization by pausing the script with the [Pause script] button.
•	Step 1 includes a preliminary optimization to find the max MU/fx limits per beam. The minimum MU/fx per beam is 800 MU.
•	Each step optimization is performed in 2 sub-runs. The first one uses the .csv objectives exactly. After the first run, temporary clinical goals are added to evaluate whether the objectives have been met. For unmet objectives, their weights are increased according to the .csv. A 2nd run is then started. The script is paused after this 2nd optimization. Script cannot be paused during optimization.
•	The extent of the dose grid along the sup-inf direction (z) is limited to slices that include a relevant structure (excluding Skin) or within 5 cm of the PTVm_2000
•	Final dose calculation: reset the dose grid to the default size. Perform final dose calculation and then set the prescription level to match the currently achieved isodose level at the ICRU point.
