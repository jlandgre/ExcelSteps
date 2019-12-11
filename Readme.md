**Overview**<br/>
ExcelSteps is a VBA add-in making spreadsheet models easy to author and maintain. It also has tremendous utility for formatting tables exported from databases and packages such as Pandas --allowing creation of easy-to-understand and nicely-formatted reports from downloaded data. It can automate both refreshing complex simulation formulas to ensure their accuracy and also can take care of formatting "janitor" work on tables. The approach brings together the following concepts that are not commonly integrated in Excel usage in enterprises large and small:

* If Excel tables are in simple rows/columns format, columns can all be automatically named by a VBA macro that sweeps through the table and does that and other useful, refreshing of the table.
* If columns are named, formulas can be symbolic within and across tables in a given workbook (e.g. Excel's naming capability anticipates Pandas DataFrames and similar, modern Data Science architectures). For example, an Area formula can be something like, `Area = (Length * Width)/2` instead of the more typical but hard to decipher `C27 = (A27 * B27)/2`
* If formulas are symbolic, they are just pieces of text that can be parked on a standard, "ExcelSteps" sheet that functions like an input deck listing tasks to be performed on a table. This makes it easy to eliminate risk of corrupt formulas in calculated columns.  It also makes it easy to add formulas into newly-added data rows.
* If formulas can be refreshed from an ExcelSteps sheet, there are numerous other routine formatting and data cleanup tasks that can be be performed whenever the workbook is refreshed. Examples include applying a drop-down list (Excel Data Validation) to a column and setting the columns' widths and number formats.
* If formulas can be refreshed it's safe to build large simulations and collaborative business-platform management models with linked tables in a single workbook. New data can be simply inserted in blank, table rows without worrying about row order, which can also be specified as a sort step in ExcelSteps.

**History of ExcelSteps Approach**<br/>
Excel is a brilliant piece of software that anticipates many aspects of the current open-source revolution. It is wonderfully (and miserably?) entrenched in enterprises large andd small --perhaps due to both cultural inertia and undeniable UI advantages that have proven difficult to replicate in any alternate software approach. Excel has the advantage and disadvantage of allowing data to be entered and summarized anywhere on its sheets in both structured and non-structured ways.

ExcelSteps works on any rows/columns table that is in the simple format of headers (aka column names) in row 1 and data in rows 2+. It overcomes common limitations of spreadsheet models making it easy and reliable to build models that link multiple worksheets. Without such an approach. models performing critical calculations easily become an error-prone "house of cards" especially if they are multi-user. This should be familiar to anyone who has ever worked in an Excel-using organization.  Here are some typical issues in the absence of an ExcelSteps approach:
* As models grow, it becomes difficult to know (like to really be confident in mission-critical calculations!) that column formulas are still correct in all rows. This is especially problematic with complex formulas involving multi-sheet VLOOKUPs and other advanced Excel functions.
* As models grow, users typically apply formatting such as merged cells to attempt to enhance the viewability of the data. However, this makes it excruciating to insert new data (and often to even know where to insert new data). At a minimum, it becomes necessary to reformat the table after inserting.
* In cases where data tables were being downloaded and cleaned in Excel, it is difficult to repeatably perform the same "recipe" of cleanup and formatting steps every time (especially when multiple users are involved). In spite of this, enterprise users settle into performing the rote steps whenever they receive data from a given source.<br/>

The ExcelSteps approach was originally integrated (at least by the original author) in 2011. It is likely that Excel's creators envisioned such an approach and there are dozens of online courses that hint at it by teaching naming and other features. However, to-date at least, the author has (surprisingly?) never run across code to put the pieces together. The first version of a VBA recipe or steps macro was actually written in the context of 2011 non-profit work --cleaning data related to a public school levy election. There, data-management volunteers needed to reliably and collaboratively clean data files downloaded from the local elections board. With no conceptual modification, the Excel automation approach from that was subsequently used for diverse applications in the author's enterprise R&D day job --both for individual model-building and in teams.<br/><br/>
 The current ExcelSteps code was created from scratch as an open-source VBA add-in in Fall 2019. The 2019 ground-up approach was based on a "cascaded function", VBA architecture. This made it possible (for the first time) to include detailed error checking and reporting. And in that approach, error conditions encountered during a step are passed back to the calling function as a non-zero result and can then be dealt with appropriately by the calling function. This allows a user-friendly error-flagging approach of placing a descriptive comment on the offending step's instruction row.<br/><br/>
J.D. Landgrebe<br/>
December 10, 2019

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
