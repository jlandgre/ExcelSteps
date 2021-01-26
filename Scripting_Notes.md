### Excel Steps Scripting Notes

This document contains raw notes pertaining to code base development.  Each day's notes are in reverse chronological "rolling scroll" format. --JDL

#### 1/26/21
* [**Complete**] Finish initial validation of Lite Scenario Model (with and without headers, as calculator and multi-column model)
* [**Complete**] To facilitate validation, create separate subroutines for refreshing scenario models on each wroksheet in Scenario Model.xlsm
* [**Complete**] Modify to get appropriate handling of header row, and scenario name row (does not need to be Model Name in row 2)
* [**Complete**] Add Cells.ClearFormats as precursor to refresh to fully check that formatting occurs properly
* [**Complete**] Pull formatting out of RefreshScenarioModel() and move to clsScenModel Format method
* **Once above steps are complete, integrate back into modScenario in add-in, test there and commit**

> ##### Scenario Model Name and HomeCell row:
> User can specify model name when provisioning clsScenModel.  This is used to name the model column if the model is a calculator (single-column) and is used as range name prefix if that is selected. If not user-specified, default model name is based on the worksheet name, clsScenModel.sht.

>If the model is single-column aka Calculator, homecell row is not a reserved space, and user's model can begin there. If multi-column model, model column names are specified by string in model homecell row. In this case, that row is a reserved space (Refresh will overwrite header and  checks whether scenario column values are appropriate as Excel range names.

#### 1/25/21

Code approach is to copy modScenario module from XLSteps.xlam and work on it in a standalone workbook

* [**Complete** - finish validation] Create "Lite" version of model that does not have Grp, Subgrp, NumFmt and Formula header columns --just Description, VarNames, Units.  Lite version can access ExcelSheets worksheet for formatting info and formulas
* [**Complete**] Create a tblExcelSheets Class to name column ranges for that sheet.  
 * This approach is amenable to merging Excel Sheets with ColInfo and having a single table with metadata about variables including how to format them
 * For now, can use the Excel Steps (rows/columns refresh) Insert formula column to hold formula text for Scenario Models also
* Additional capability for Excel Steps would be for it to recognize and remove a "tag" prefix on formulas and number formats which can get mis-interpreted in Excel (and especially in *.CSV files that have no formatting). These items can either be represented as text such as "0.00" for a format string or they can be mis-interpreted as the number 0 by Excel if not formatted as text.  Perhaps use a "~" as a prefix character ignored by the VBA code

Note - need to check case sensitivity in ExcelSteps lookup.  Does Sheet and Column name case matter?

1/22/21
* Convert RefreshScenarioModel function to using Class Instance as its basis
* clsScenModel Class holds description of the Scenario model including its position within a workbook and sheet
