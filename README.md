Excel Column Hiding and Protection Project

Overview
This Excel VBA project automates the process of hiding columns in a worksheet based on dates. It checks the dates in row 6 of the "Sheet1" sheet and hides columns where the date is earlier than the current date. The script also manages worksheet protection, ensuring rows 8 through 45 remain editable while the rest of the sheet is locked.

The code runs automatically when the workbook is opened.

Features
- Hides columns based on a date comparison (dates before today are hidden).
- Protects the worksheet with a password (password), leaving rows 8–45 editable.
- Automatically triggers on workbook open.

Repository Structure
- Module1: Contains the main logic for hiding columns (HideColumnsBasedOnDate) and a workbook open event handler.
- ThisWorkbook: Contains an additional workbook open event handler to ensure proper initialization and protection.

Prerequisites
- Microsoft Excel with VBA enabled.
- The workbook must contain a sheet named "Sheet1".
- Dates to compare must be in row 6 of the sheet.

Installation
1. Clone or download this repository to your local machine.
2. Open the Excel file containing your project (e.g., YourFile.xlsm).
3. Open the VBA Editor (Alt + F11) and import the code:
   Drag Module1 and ThisWorkbook into the VBA Project Explorer, or copy-paste the code into the respective modules.
4. Save the workbook as a macro-enabled file (.xlsm).

Usage
1. Open the Excel workbook.
2. The macro will automatically run, hiding columns in the "Sheet1" sheet where the date in row 6 is earlier than today.
3. Rows 8–45 remain editable; all other cells are protected with the password password.

Code Details
- Sub HideColumnsBasedOnDate: 
  --Unprotects the sheet (password: password).
  --Locks all cells, then unlocks rows 8–45.
  --Loops through dates in row 6, hiding columns with dates before today.
  --Reprotects the sheet, allowing VBA to run without user intervention.

- Workbook_Open: 
  --Triggers HideColumnsBasedOnDate when the workbook opens.
  --Ensures rows 8–45 remain editable after protection.

Customization
- Sheet Name: Update ThisWorkbook.Sheets("Sheet1") if your sheet has a different name.
- Date Threshold: Modify comparisonDate = Date in HideColumnsBasedOnDate to use a specific date (e.g., comparisonDate = #12/31/2024#).
- Password: Change password to a different password if needed.
- Range: Adjust ws.Range(ws.Cells(6, 1), ws.Cells(6, lastCol)) if your dates are in a different row.

Example
If row 6 contains the following dates:

| A      | B      | C      | D      |
|--------|--------|--------|--------|
| 1/1/24 | 5/1/24 | 4/1/25 | 6/1/25 |

On March 6, 2025, columns A and B will be hidden, while C and D remain visible.

License
This project is open-source and available under the MIT License (LICENSE). Feel free to modify and distribute it as needed.

Contributing
1. Fork this repository.
2. Create a new branch (git checkout -b feature-branch).
3. Commit your changes (git commit -m "Add feature").
4. Push to the branch (git push origin feature-branch).
5. Open a pull request.

Contact
For questions or suggestions, feel free to open an issue in this repository.

