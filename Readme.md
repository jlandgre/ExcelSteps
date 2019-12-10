ExcelSteps is an add-in that brings together several Excel features to make spreadsheet models easy to author and maintain.  It also has tremendous utility as a way to format tables exported from other packages such as Pandas.  ExcelSteps is based on bringing together the following, disparate concepts into a whole that proves valuable in many situations:

* If Excel tables are in simple rows/columns format, columns can all be automatically named by a macro that sweeps through the table and does that and other useful, refreshing of the table.
*  If columns are named, formulas can  be symbolic within and across tables in a given workbook (e.g. Excel's naming capability anticipates Pandas DataFrames and similar, modern Data Science architectures), so an area formula can be something like, `Area = (Length * Width)/2` instead of the more typical but hard to decipher and maintain `C27 = (A27 * B27)/2`
*  If formulas are symbolic, they are just pieces of text that can be parked on a standard, "ExcelTasks" sheet that functions like an input deck listing tasks to be performed on a table.  This makes it easy to refresh tables to eliminate risk of corrupt formulas in calculated columns and/or to add the formulas to newly-added data rows.
*  If formulas can be refreshed from an ExcelTasks sheet, there are numerous other formatting  and data cleanup tasks that can also be listed as refresh instructions.  Examples include mundane things like applying a drop-down list (Excel Data Validation) to a column and setting the columns' widths and number formats.
*  If formulas can be refreshed, they can't be [permanently] broken, and it's safe to build large simulations and collaborative business-platform management models with linked tables in a single workbook. New data can be simply inserted in blank, table rows without worrying about row order. Refresh functionality then takes care of adding column formulas, formatting and sorting appropriately.

J.D. Landgrebe<br/>
December 10, 2019

Example<br/>
As a basic example, this repository contains the file `Right Triangles Example.xlsx`.  To run the example, first install the Add-in file, `XLSteps.xlam` through Excel's Options menu --> Add-ins --> Go --> Browse.
1.  Open `Right Triangles Example.xlsx`.  Note that it is a simple, unformatted data table representing the lengths of sides A and B in a few right triangles.
<br/><br/><img src=Assets/Triangles1.png alt="Unmodified table" width=250><br/>

2.  Choose `Refresh Excel Tables` from the Add-ins ribbon's ExcelSteps menu.
<br/><br/><img src=Assets/AddinsMenu.png alt="Excel Add-ins Ribbon" width=400><br/>

3.  In the dialog box, select the Refresh and Replace checkboxes next to the `Triangles` worksheet name.  Click OK
<br/><br/><img src=Assets/Triangles2.png alt="Excel Steps Refresh Dialog" width=300><br/>

4.  The refreshed table now has some formatting, but note that the refresh also created Names for the individual columns, the overall table and its header row.
<br/><br/><img src=Assets/Triangles3.png alt="Refreshed Table" width=250>  <img src=Assets/NameManager.png alt="Name Manager Post-Refresh" width=400><br/>

5. Notice also that the initial refresh added a blank Excel Steps worksheet.  Enter a row as shown on this sheet and repeat the refresh.
<br/><br/><img src=Assets/Excel_steps_blank.png alt="Unmodified table" width=600>
<br/><br/>
 <img src=Assets/Excel_Steps2.png alt="Unmodified table" width=625><br/>

6. ExcelSteps Inserts a new column, `Side_C` with a header comment and formatting as specified in the Excel Steps row.
<br/><br/>
 <img src=Assets/Triangles4.png alt="Unmodified table" width=300><br/>
