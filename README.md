# UPV_DisplaySet_VBA
Written by Prattik Rodgers
📘 UPV Cable Package Generator – User Guide
________________________________________
🧠 What this tool does:
This macro-driven Excel tool allows you to generate UPV-ready routing packages from cable data by:
•	Verifying cable names against the master cable schedule
•	Extracting and formatting routing information
•	Packaging the data into a structured format for import into the UPV model
________________________________________
🚦 Three Ways to Use the Tool
🔹 1. Cable Name Entry (Manual List Mode)
Best for: You know the exact cable names and want to generate a package for them.
✅ Steps:
1.	Go to the Cable Name List sheet
2.	Enter your cable names starting in cell B3 (one per row)
3.	Press Alt + F8
4.	Select and run A1StartCableListUPV
🛠 What Happens:
•	Each cable name is checked against CableScheduleData!D:D
•	If any cable name is not found, you'll get:
o	A red highlight on the invalid cell
o	A message box with all missing names
o	The macro will stop until all errors are fixed
•	If everything matches, the routing data is gathered and the package is auto-generated
💾 Final Step:
•	Go to the RecieveUPVFileHere sheet
•	Copy the sheet into a new .xlsx workbook
•	Save and upload it into the model system
•	🎉 You're done!
________________________________________
🔹 2. Cable Schedule Filter Mode (Auto Pull from Filtered Data)
Best for: You’ve already filtered the full schedule and want to generate a package from visible rows.
✅ Steps:
1.	Go to the CableScheduleData sheet
2.	Apply your filter (by area, status, or anything else)
3.	Press Alt + F8
4.	Run A1CableScheduleSort
🛠 What Happens:
•	The macro automatically:
o	Reads only visible (filtered) cable names
o	Transfers them into Cable Name List
o	Triggers A1StartCableListUPV to run automatically
•	The rest of the process is identical to Option 1
💾 Final Step:
•	Copy the RecieveUPVFileHere sheet to a clean workbook
•	Save and upload it into the model system
________________________________________
🔹 3. Area Completion Mode (Advanced Filter Based on % Complete)
Best for: You want a package based on completion status (0%, partial, full) within a specific area.
✅ Steps:
1.	In In8:
o	Export a Quantity Tracking report
o	Ensure it includes:
	Area column (only numeric, no trailing letters)
	Quantity, To Date Quantity, and % Complete
	Name column filtered to entries with “Cable Pull”
2.	Create a Data Export from In8
o	Download and open the file
o	Enable editing
o	Drag the Template sheet into the UPV Generator workbook
3.	Go to “Input your area here” sheet
o	In cell F2, enter the Area Code (must be numeric only — e.g., 221 instead of 221U)
o	In G2, select one of the following filter types:
	All — includes all rows
	Partial — includes 0% < % Complete < 95%
	Complete — includes 95% ≤ % Complete ≤ 100%
	Zeroed — includes only rows with 0% complete
4.	Press Alt + F8
5.	Run A1StartAreaUPV
🛠 What Happens:
•	The macro filters the Template sheet based on area code and completion status
•	Excludes any rows with ELP, LPP, or RPP in the Name column
•	Validates cable names against the master schedule
•	Packages the data and builds the UPV output
💾 Final Step:
•	Copy the RecieveUPVFileHere sheet into a new .xlsx file
•	Save and upload that file to the model
________________________________________
🧠 Tips & Notes
•	Highlighted cable names in red mean they weren’t found — fix them before continuing.
•	Make sure macros are enabled before using the tool.
•	Always save your workbook before running macros to avoid losing data.
________________________________________
✅ Final Output:
Once the macro completes:
•	Your data will be fully formatted in RecieveUPVFileHere
•	Ready for copy/paste or upload into your system
