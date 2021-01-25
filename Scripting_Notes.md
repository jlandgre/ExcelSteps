### Excel Steps Scripting Notes

This document contains raw notes about code base development.  These are in reverse chronological "rolling scroll" format.

JDL

1/25/21
* [Complete - finish validation] Create "Lite" version of model that does not have Grp, Subgrp, NumFmt and Formula header columns --just Description, VarNames, Units.  Lite version can access ExcelSheets worksheet for formatting info and formulas
* [complete] Create a tblExcelSheets Class to name column ranges for that sheet.  
 * This approach is amenable to merging Excel Sheets with ColInfo and having a single table with metadata about variables including how to format them
 * For now, can use the Excel Steps (rows/columns refresh) Insert formula column to hold formula text for Scenario Models also
* Additional capability for Excel Steps would be for it to recognize and remove a "tag" prefix on formulas and number formats which can get mis-interpreted in Excel (and especially in *.CSV files that have no formatting). These items can either be represented as text such as "0.00" for a format string or they can be mis-interpreted as the number 0 by Excel if not formatted as text.  Perhaps use a "~" as a prefix character ignored by the VBA code

Note - need to check case sensitivity in ExcelSteps lookup.  Does Sheet and Column name case matter?

1/22/21
* Convert RefreshScenarioModel function to using Class Instance as its basis
* clsScenModel Class holds description of the Scenario model including its position within a workbook and sheet
