**Overview**<br/>
ExcelSteps is a Microsoft Excel VBA add-in for curating and managing calculation models. It defines rows/columns table objects and/or “Scenario Model” columns-by-rows objects within a workbook. It makes calculation formulas and formatting instructions refreshable. It creates named ranges for each variable in the model, and it names ranges for the table and Scenario Model objects to aid cross-object lookups and references. 

ExcelSteps is currently Windows-only. It can be used in two modes. There is a user-facing menu for refreshing models. It also has an API code design for programmatically defining and managing custom table and Scenario Model objects. This provides within-workbook location flexibility and a full range of configuration options. On the Excel (user-facing) side, model wayfinding is based on named ranges and on Control-F searching of standard variable name locations. Additionally, the VBA Class objects include lookup methods for accessing values and ranges programmatically --modeled after Python Pandas .loc() functionality. An ExcelSteps “recipe” sheet holds formatting and calculation formula instructions for the model. This houses these items in a central location for easy editing and validation.
</br>
<center>ExcelSteps Addin menu (MS Excel for Windows)</center>
<img src=Assets/models1_menu.png alt="ExcelSteps Addin menu"
    style="display:block;
            float:none;
            margin-left:auto;
            margin-right:auto;
            width:200px;">
</br>
<center>Refreshed Rows/Columns Model</center>
<img src=Assets/Triangles4.png alt="Example Rows/Columns model"
    style="display:block;
            float:none;
            margin-left:auto;
            margin-right:auto;
            width:400px;">
</br>

As an alternative to rows/columns tables, so-called Scenario Models serve two purposes. First, they provide an alternate columns by rows aesthetic allowing a model to flow from left-to-right or to consist of side-by-side scenarios. This is sometimes an intuitive calculation model aesthetic, and it is a way to put a calculation model into the hands of non-coders such as business executives. Because of their standard format consisting of standard left-side header columns and Row 2 scenario names, default (aka menu-refreshed) scenario models can be transposed into rows/columns using the addin's **Parse Scenario Model** menu. This is convenient for flipping a model for graphing and portability to other applications. In consulting projects, a custom-configured single-column Scenario Model is an intuitive way to display and let users interact with single-valued model parameters --often adjacent to a rows/columns table containing related data.

</br>
<center>Refreshed "Default" Multi-column Scenario Model</center>
<img src=Assets/models1.png alt="Example Scenario Models"
    style="display:block;
            float:none;
            margin-left:auto;
            margin-right:auto;
            width:700px;">
</br>
<center> Parsed (aka Transposed) Scenario Model</center>
<img src=Assets/models1_parsed.png alt="Parsed Scenario Model"
    style="display:block;
            float:none;
            margin-left:auto;
            margin-right:auto;
            width:300px;">
</br>
<center>Programmatically-Configured Scenario Model</center>
<img src=Assets/models7.png alt="Example Scenario Models"
    style="display:block;
            float:none;
            margin-left:auto;
            margin-right:auto;
            width:750px;">
</br>

The vision for ExcelSteps was to base it on traditional Excel features --making calculation models accessible and usable by a broader range of users than would be possible using “new Excel” features. These are more focused on analytics and Data Science professionals and are not as thoughtfully designed for calculation model usage and validation. 

ExcelSteps is based on the following logic for rows/columns models –with analogous thinking for the transposed Scenario Model format:
* If tables and Scenario models are in standard formats, they can be automatically named by a VBA macro that sweeps through model's variables and key ranges such as a table's header row or a Scenario Model's variable names column.
* If ranges holding variable values are named, formulas can be symbolic within and across objects in a workbook (e.g. Excel's naming capability anticipates Pandas DataFrames and similar, modern Data Science architectures). For example, an Area formula for a right triangle can be **Area = (Length * Width)/2** instead of the harder-to-decipher **C27 = (A27 * B27)/2**
* If formulas are symbolic, they are just pieces of text that can be stored on a standard, "ExcelSteps" recipe sheet functioning like an input deck listing tasks to perform. This eliminates risk of corrupt formulas in calculated columns. It also makes it easy to add new formulas as additional recipe rows.
* If formulas can be refreshed from an ExcelSteps sheet, there are numerous routine formatting and data cleanup tasks that can be refreshed. Examples include applying a drop-down list (Excel Data Validation) to a column and setting columns' widths and number formats.
* If formulas can be refreshed it's safe to build large simulations and collaborative business-platform management models with linked tables and Scenario Models in a single workbook.

Although we also love the Mac platform, ExcelSteps is currently Windows only as of July 2023. ExcelSteps was based on primordial insights dating back to the early 2010's. Its initial development was in January 2020 with continual enhancements from usage in consulting work. As of July 2023 the Scenario model (mdlScenario) and table (tblRowsCols) VBA objects are validated by test suites in the **modTests_mdl**, and **modTests_tbl** modules of **ExcelSteps_Validation.xlsm**. Despite the validation, the code is not warranted against errors. It is recommended that you perform your own, thorough application-specific validation if you use ExcelSteps in your work.

J.D. Landgrebe
Data Delve Engineer LLC


**Example**<br/>
As a basic example, this repository contains the file `Right Triangles Example.xlsx`. To run the example, first install the Add-in file, `XLSteps.xlam` through Excel's Options menu --> Add-ins --> Go --> Browse.
1. Open `Right Triangles Example.xlsx`. Note that it is an unformatted data table representing the lengths of sides A and B in a few right triangles.
<br/><br/><img src=Assets/Triangles1.png alt="Unmodified table" width=250><br/>

2. Choose `Refresh Excel Tables` from the Add-ins ribbon's ExcelSteps menu.
<br/><br/><img src=Assets/models1_menu.png alt="Excel Add-ins Ribbon" width=200><br/>

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

 **January 2020 Background on ExcelSteps and History of its Approach**<br/>
 Microsoft Excel's ground-breaking design dates back to the 1980s. It anticipates many aspects of the current open-source revolution. It is entrenched in enterprises large and small --due to both cultural inertia and undeniable UI advantages difficult to replicate in alternate software designs. Excel has the wonderful advantage and terrible disadvantage of allowing data to be entered and summarized anywhere on its sheets in both structured and non-structured ways. It also includes features like Merged Cells that encourage users to format tables in aesthetically pleasing ways but which make tables difficult to maintain and edit.

 It is likely that Excel's creators envisioned the ExcelSteps approach, but, to my knowledge, they never articulated it clearly or implemented it with upgraded features. Instead, Microsoft got distracted by "New Excel" features like Tables, PowerQuery and DAX for analytics pros. For the average user, these represent a need to learn a second language though, and that's not adopted widely in non-analytics functions. 
 
 There are dozens of online courses that hint at the ExcelSteps approach by teaching naming and other features. However, it has never become commonplace to put these pieces together in enterprises and research efforts large and small. ExcelSteps automates refreshing any rows/columns table that is in the simple format of headers (aka column names) in Row 1 and data in Rows 2+. It makes it easy and reliable to build models linking multiple worksheets. Without such an approach, models performing critical calculations easily become an error-prone "house of cards" especially if they are used by multiple people. This should be familiar to anyone who has ever worked in an Excel-using organization. Here are some typical issues that crop up:
 * Spreadsheet models get built as unstructured blobs of calculations rather than "pure" rows/columns tables. The blobs mix data and summaries of the data. These are difficult to error-check, link across tables and export to stats and visualization packages.
 * As models grow, it becomes difficult to know (like to really be confident in mission-critical calculations!) that column formulas are still correct in all rows. This is especially problematic with complex formulas involving multi-sheet VLOOKUPs and other advanced Excel functions.
 * As models grow, users typically apply formatting such as merged cells to enhance viewability. However, this makes it excruciating to insert new data. It often becomes difficult to even know where to insert new data. At a minimum, it becomes necessary to reformat the table after making even simple additions.
 * When data tables are being downloaded and cleaned in Excel, it is difficult to repeatably perform the same "recipe" of cleanup and formatting steps every time (especially when multiple users are involved). In spite of this, users often resign themselves to performing these data cleanup steps manually whenever they receive data from a given source.<br/>

