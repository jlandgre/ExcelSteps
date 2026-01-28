# Skill: VS Code Environment Setup for VBA Projects with ExcelSteps

## Overview
This document describes how to set up a VS Code workspace for developing VBA projects that use the ExcelSteps add-in. The setup involves managing multiple Excel files across different folders, with separate PowerShell terminals for each file to enable real-time code synchronization via xlwings.

## Project Structure

A typical project has this folder organization:
```
Projects/
 project_name/                    # Master project folder
    .github/                     # Project-specific GitHub settings and skills
    src/                         # Project source code
       project_name.xlsm       # Main project workbook
    tests/                       # Test suite
       tests_project_name.xlsm # Test workbook
    project_name.code-workspace # VS Code workspace file
 Excel_Steps/                     # ExcelSteps add-in (shared across projects)
     .github/                     # ExcelSteps GitHub settings and skills
     src/
         XLSteps.xlam            # ExcelSteps add-in file
```
A special case of this is working on the XLSteps.xlam add-in itself. In this case, the Excel_Steps folder contains the src and tests subfolder and there are only two files involved (XLSteps.xlam and tests_XLSteps.xlsm).

## Files Being Edited

When working on a project, you typically edit three Excel files simultaneously:

1. **Project File**: `Projects/project_name/src/project_name.xlsm`
2. **Test File**: `Projects/project_name/tests/tests_project_name.xlsm`
3. **ExcelSteps Add-in**: `Projects/Excel_Steps/src/XLSteps.xlam`

Each file requires its own PowerShell terminal running `xlwings vba edit` to sync VBA code between the Excel workbook and file system.

## Initial Setup

### 1. Open Parent Project Folder
**IMPORTANT:** To ensure the `.github` folder is accessible to Copilot:
1. Open the parent project folder in VS Code: **File  Open Folder**
2. Select the master project folder: `Projects/project_name/`
3. This makes the `.github/skills` folder and Copilot instructions discoverable

### 2. Create or Open Workspace
1. If a workspace file exists: **File  Open Workspace from File**  Select `project_name.code-workspace`
2. If it doesn't exist, you'll add folders next and save the workspace

### 3. Configure Workspace Folders
Add the following folders to your workspace:
- `Projects/project_name/src`
- `Projects/project_name/tests`
- `Projects/Excel_Steps/src` (if editing the add-in)

**VS Code:** File  Add Folder to Workspace  Select each folder

### 4. Save Workspace (if new)
If you created a new workspace:
- **File  Save Workspace As**
- Save as `project_name.code-workspace` in the master folder `Projects/project_name/`

### 5. Verify Project Structure
Ensure the following are present:
- `.github/` folder exists in `Projects/project_name/`
- Workspace file saved as `project_name.code-workspace` in master folder
- All Excel files are present in their respective locations

## Setting Up PowerShell Terminals

### 1. Create Terminal for Project File
1. Open new PowerShell terminal: **Terminal  New Terminal**
2. Navigate to project src folder:
   ```powershell
   cd "Z:\path\to\Projects\project_name\src"
   ```
3. **IMPORTANT:** Rename terminal immediately: Right-click terminal tab  **Rename**  Enter `project_name`
   - This clarifies which file the terminal manages, even if the folder path shown is different

### 2. Create Terminal for Test File
1. Create another PowerShell terminal
2. Navigate to tests folder:
   ```powershell
   cd "Z:\path\to\Projects\project_name\tests"
   ```
3. **IMPORTANT:** Rename terminal to: `tests_project_name`
   - The terminal tab may show the previous folder initially, but the rename makes it clear

### 3. Create Terminal for ExcelSteps (if editing add-in)
1. Create another PowerShell terminal
2. Navigate to ExcelSteps src folder:
   ```powershell
   cd "Z:\path\to\Projects\Excel_Steps\src"
   ```
3. **IMPORTANT:** Rename terminal to: `XLSteps`

**NOTE:** VS Code may display the folder path of where the terminal was initially created in the terminal tab. This is why renaming is criticalthe custom name clearly identifies which file each terminal manages, regardless of what folder path is shown.

## Enabling xlwings VBA Edit

For each file you want to edit, run `xlwings vba edit` in its corresponding terminal:

### Project File Terminal (`project_name`)
```powershell
xlwings vba edit -f "project_name.xlsm"
```
When prompted, type `Y` to proceed.

### Test File Terminal (`tests_project_name`)
```powershell
xlwings vba edit -f "tests_project_name.xlsm"
```
When prompted, type `Y` to proceed.

### ExcelSteps Terminal (`XLSteps`)
```powershell
xlwings vba edit -f "XLSteps.xlam"
```
When prompted, type `Y` to proceed.

## Working with the Setup

### Active xlwings Sessions
Once `xlwings vba edit` is running in a terminal:
- You'll see: `Watching for changes in [filename] (silent mode)...(Hit Ctrl-C to stop)`
- Any changes to `.bas`, `.cls`, or `.frm` files in VS Code will sync to the Excel workbook
- Any changes in the VBA Editor will sync to the file system
- **Keep this running while working** - do not close or interrupt the terminal

### Terminal Naming Convention
Use clear, descriptive names for terminals:
- **Project file**: `project_name` or just the project name (e.g., `Dashboard`, `Forecast`)
- **Test file**: `tests_project_name` (e.g., `tests_Dashboard`, `tests_Forecast`)
- **ExcelSteps**: Always `XLSteps`

**Why this matters:** The terminal's displayed folder path may not reflect the actual working directory after navigation. Custom terminal names provide clarity about which file each terminal manages.

### Opening Excel Files
To open the Excel files in the project:
1. Navigate to the appropriate folder in VS Code Explorer
2. Right-click the Excel file  **Reveal in File Explorer**
3. Double-click to open in Excel

Or use PowerShell:
```powershell
# From project terminal
Start-Process "project_name.xlsm"

# From tests terminal
Start-Process "tests_project_name.xlsm"

# From XLSteps terminal
Start-Process "XLSteps.xlam"
```

## Critical Workflow Rules

### Before Making Code Changes
**Always verify** that `xlwings vba edit` is running for any file you plan to modify:
1. Check terminal for the message: `Watching for changes in [filename]...`
2. If not running, start it before making code changes
3. This ensures changes sync properly between VS Code and Excel

### Syncing Changes
- **From Excel to VS Code**: Use `xlwings vba import` in the terminal (one-time import)
- **Continuous sync**: Keep `xlwings vba edit` running (watches for changes both ways)

## Example: Setting Up Excel_Steps Project

For the Excel_Steps project itself (working on the add-in and its tests):

### Step 1: Open Parent Folder
1. **File  Open Folder**
2. Select `Z:\Users\j.d.landgrebe\Box Sync\Projects\Excel_Steps`
3. This makes `.github/` accessible

### Step 2: Set Up Workspace Folders
Add folders:
- `Z:\Users\j.d.landgrebe\Box Sync\Projects\Excel_Steps\src`
- `Z:\Users\j.d.landgrebe\Box Sync\Projects\Excel_Steps\tests`

### Step 3: Set Up Terminals
```powershell
# Terminal 1: Create, navigate, and RENAME to "XLSteps"
cd "Z:\Users\j.d.landgrebe\Box Sync\Projects\Excel_Steps\src"
xlwings vba edit -f "XLSteps.xlam"

# Terminal 2: Create, navigate, and RENAME to "tests_XLSteps"
cd "Z:\Users\j.d.landgrebe\Box Sync\Projects\Excel_Steps\tests"
xlwings vba edit -f "tests_XLSteps.xlsm"
```

Workspace file: `Excel_Steps.code-workspace` in `Projects/Excel_Steps/`

## Troubleshooting

### Terminal Shows Exit Code 1
- The file may not exist in that location
- Check the file name and path
- Ensure you're in the correct directory

### Changes Not Syncing
- Verify `xlwings vba edit` is still running (not stopped or exited)
- Check for error messages in the terminal
- Try stopping (Ctrl-C) and restarting `xlwings vba edit`

### Terminal Display Confusion
- If a terminal shows a different folder than expected in the tab, check the custom name you assigned
- The custom terminal name (e.g., `XLSteps`, `tests_XLSteps`) is the authoritative indicator
- You can verify the actual directory with `pwd` or `Get-Location` in the terminal

### .github Folder Not Found
- Make sure you opened the parent project folder (`Projects/project_name/`) not just added workspace folders
- Use **File  Open Folder** to open the parent directory
- This is required for Copilot to discover the `.github/skills` folder

### Multiple Projects
When switching between projects, you can:
- Stop all `xlwings vba edit` sessions (Ctrl-C in each terminal)
- Open the new project's workspace file
- Set up terminals and restart `xlwings vba edit` for the new project

## Best Practices

1. **Open parent folder first** - Ensures `.github` is accessible before adding workspace folders
2. **Always rename terminals immediately** - Makes it clear which file each terminal manages, regardless of displayed path
3. **Keep terminals visible** - Helps you verify xlwings is running before code changes
4. **One workspace per project** - Don't mix multiple projects in the same workspace
5. **Save workspace regularly** - Preserves your terminal setup and folder configuration
6. **Close Excel before major VBA imports** - Prevents file locking issues
