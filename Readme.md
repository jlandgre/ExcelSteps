**Overview**<br/>
ExcelSteps is a VBA add-in that makes spreadsheet models easy to author and maintain. It also formats tables exported from databases and packages such as Pandas --allowing creation of reports. It can automate both refreshing complex simulation formulas to ensure their accuracy and also can take care of "janitor" work to keep tables formatted as they are updated. Here is a 5-minute video showing a demo: https://vimeo.com/manage/399719128/general. The approach brings together the following concepts:

* If Excel tables are in simple rows/columns format, columns can be automatically named by a VBA macro that sweeps through the table and does that and other useful, refreshing.
* If columns are named, formulas can be symbolic within and across tables in a given workbook (e.g. Excel's naming capability anticipates Pandas DataFrames and similar, modern Data Science architectures). For example, an Area formula for a right triangle can be something like, `Area = (Length * Width)/2` instead of the more typical but harder to decipher `C27 = (A27 * B27)/2`
* If formulas are symbolic, they are just pieces of text that can be parked on a standard, "ExcelSteps" recipe sheet that functions like an input deck listing tasks to be performed on a table. This eliminates risk of corrupt formulas in calculated columns.  It also makes it easy to add formulas into newly-added data rows.
* If formulas can be refreshed from an ExcelSteps sheet, there are numerous other routine formatting and data cleanup tasks that can be be performed whenever the workbook is refreshed. Examples include applying a drop-down list (Excel Data Validation) to a column and setting the columns' widths and number formats.
* If formulas can be refreshed it's safe to build large simulations and collaborative business-platform management models with linked tables in a single workbook. New data can be simply inserted in blank, table rows without worrying about row order, which can also be specified as a sort step in ExcelSteps.

**Notes on Current Release**<br/>
January 15, 2021:  Debugged scenario model (using Scenario Model.xlsm test cases) to get it working better with a header row; prep for "refactoring" original, quickly-constructed modScenario.  Plan to create a clsScenarioModel class to hold a model's attributes including its location in a workbook.  This will make it easier to work with applications that have multiple Scenario models in a project workbook.  It also will greatly simplify and clean up modScenario based on reapplying learning from a Class Instance approach with rows/columns models

January 11, 2021: Significant updates to enable use as an API by calling RefreshRC() directly from VBA instead of from the ExcelSteps menu and dialog box.  Also updated such that most recipe commands work with "floating" tables that not homed to cell A1 (header) and cell A2 (data rows).  Row/Column refresh accepts an arbitrary header range as input (rngheader argument to RefreshRC subroutine) and with arbitrary data row range (rngData argument).  The menu instructions in in Validation Tests_API.xlsm have been verified with this version.  

March 16, 2020: Updated Triangles.xlsx case study to account for an under-publicized, significant update to Excel itself.  In January 2020, Microsoft pushed an update to Windows Excel that now requires column formula references to include an '@' character as a prefix (e.g. "=@Side_a" refers to the row-specific cell in the Side_a column --just that one cell in the same row as the formula).  While the '@' character has long worked this way, it has also long been optional. Historically, the formula "=Side_a" would produce the same result as "=@Side_a".  As of the MS update, the range name without an '@' character now refers exclusively to the entire column.  If row-specific formulas are used without the '@' character, Excel displays a '#SPILL' error.

March 16, 2020: Fixed bug with Freeze Row 1 recipe command ; updated code to handle case where ExcelSteps worksheet is itself exported as a Pandas DataFrame using the Pandas ExcelWriter.  Pandas exports tables including their index column.  ExcelSteps now recognizes this and deletes the index column from the recipe when the workbook is initially refreshed.

January 14, 2020: added settings persistence so that the refresh dialog box's settings would not be cleared after every refresh.  This was done by adding a hidden settings sheet to the workbook (set to xlSheetVeryHidden .visible status in VBA) to remember settings from the previous refresh.

As of December 2019, ExcelSteps is validated per the documentation in the repository. Two of the Steps options can perform multiple actions. Col_Insert inserts (or refreshes an existing version of) a column with the option to include a calculation formula.  Col_Insert can also apply number formatting, a header comment and column width to the column. The example below shows this. The Col_Format step allows specifying number formatting, header comment text and width in a single step. Additionally, the release contains steps to sort the table by multiple keys (Tbl_Sort) and a Col_CondFormat step to discourage use of merged cells. It applies borders around groups of repeated values in a column's rows.  Repeat values are conditionally formatted in white text and therefore not visible --giving an appearance similar to merged cells.<br/>

**Next Steps**<br/>
* Add a DeleteColumn recipe step
* Add capability to automatically name blank columns that are within a table's header range.  This will allow ExcelSteps to handle the use case where an imported table includes unnamed columns (such as occurs with a Pandas DataFrame whose index is unnamed).  ExcelSteps could a) name these with an indexed name such as "Unnamed_0", Unnamed_1" etc. and b) optionally delete such columns if unneeded.
* Augment ExcelSteps with additional steps such as column Rename and Replace and to add additional table formatting steps for freezing the first row or selected columns in a table.
* Add automatic range naming on Lists worksheet to make it easy to park dropdown list values there
* Automate re-validation by writing code that walks through all checks and compares results versus a verified standard file
* Consider eliminating individual formatting steps in ExcelSteps --in favor of simply using Col_Format for tasks like adding a header comment

**Example**<br/>
As a basic example, this repository contains the file `Right Triangles Example.xlsx`. To run the example, first install the Add-in file, `XLSteps.xlam` through Excel's Options menu --> Add-ins --> Go --> Browse.
1. Open `Right Triangles Example.xlsx`. Note that it is a simple, unformatted data table representing the lengths of sides A and B in a few right triangles.
<br/><br/><img src=Assets/Triangles1.png alt="Unmodified table" width=250><br/>

2. Choose `Refresh Excel Tables` from the Add-ins ribbon's ExcelSteps menu.
<br/><br/><img src=Assets/AddinsMenu.png alt="Excel Add-ins Ribbon" width=400><br/>

3. In the dialog box, select the Refresh and Replace checkboxes next to the `Triangles` worksheet name. Click OK
<br/><br/><img src=Assets/Triangles2.png alt="Excel Steps Refresh Dialog" width=300><br/>

4. The refreshed table now has some formatting, but note that the refresh also created Names for the individual columns, the overall table and its header row.
<br/><br/><img src=Assets/Triangles3.png alt="Refreshed Table" width=250> <img src=Assets/NameManager.png alt="Name Manager Post-Refresh" width=400><br/>

5. Notice also that the initial refresh added a blank Excel Steps worksheet. Enter a row as shown on this sheet and repeat the refresh.
<br/><br/><img src=Assets/Excel_steps_blank.png alt="Unmodified table" width=600>
<br/><br/>
 <img src=Assets/Excel_Steps2.png alt="Unmodified table" width=625><br/>

6. ExcelSteps Inserts a new column, `Side_C` with a header comment and formatting as specified in the Excel Steps row.
<br/><br/>
 <img src=Assets/Triangles4.png alt="Unmodified table" width=300><br/>

 **Background on ExcelSteps and History of its Approach**<br/>
 Excel's design anticipates many aspects of the current open-source revolution. It is wonderfully entrenched in enterprises large and small --due to both cultural inertia and undeniable UI advantages difficult to replicate in alternate software designs. Excel has the wonderful advantage and terrible disadvantage of allowing data to be entered and summarized anywhere on its sheets in both structured and non-structured ways.  It also includes features like Merged Cells that encourage users to format tables in aesthetically pleasing ways but which make tables difficult to maintain and edit.

 It is likely that Excel's creators envisioned the ExcelSteps approach. There are dozens of online courses that hint at it by teaching naming and other features. However, it has  never become commonplace to put these pieces together in enterprises and research efforts large and small. ExcelSteps automates refreshing any rows/columns table that is in the simple format of headers (aka column names) in Row 1 and data in Rows 2+. It makes it easy and reliable to build models linking multiple worksheets. Without such an approach. models performing critical calculations easily become an error-prone "house of cards" especially if they are used by multiple people. This should be familiar to anyone who has ever worked in an Excel-using organization.  Here are some typical issues that crop up:
 * Spreadsheet models get built as unstructured blobs of calculations rather than "pure" rows/columns tables.  The blobs mix data and summaries of the data.  These are difficult to error-check, link across tables and export to stats and visualization packages.
 * As models grow, it becomes difficult to know (like to really be confident in mission-critical calculations!) that column formulas are still correct in all rows. This is especially problematic with complex formulas involving multi-sheet VLOOKUPs and other advanced Excel functions.
 * As models grow, users typically apply formatting such as merged cells to enhance viewability. However, this makes it excruciating to insert new data.  It often becomes difficult to even know where to insert new data. At a minimum, it becomes necessary to reformat the table after making even simple additions.
 * When data tables are being downloaded and cleaned in Excel, it is difficult to repeatably perform the same "recipe" of cleanup and formatting steps every time (especially when multiple users are involved). In spite of this, users often resign themselves to performing these data cleanup steps manually whenever they receive data from a given source.<br/>

 The ExcelSteps approach overcomes these limitations. It was originally integrated in 2011 VBA code for non-profit volunteer work wrangling data for a U.S. Ohio public school levy election. That code and its steps/recipe template subsequently proved useful to the author's enterprise R&D work at P&G. It was informally disseminated there in the form of an add-in.

 The current, open-source ExcelSteps add-in was created from the ground up in Fall 2019. The 2019 approach was based on a "cascaded function," VBA architecture. This made it possible to include detailed error checking and reporting. Error conditions encountered during a step are passed back to the calling function as a non-zero result and can then be dealt with appropriately by the calling function. This allows a user-friendly error-flagging approach of placing a descriptive comment on the offending step's instruction row.<br/><br/>

 J.D. Landgrebe<br/>
 December 10, 2019
