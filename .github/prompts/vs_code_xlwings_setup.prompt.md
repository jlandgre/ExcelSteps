---
name: VS Code XLwings Setup
description: >
  Set up VS Code terminals for VBA development with xlwings, managing multiple Excel files across folders with real-time code sync.
---
# Skill: VS Code Environment Setup for VBA Projects with ExcelSteps

## Overview
This describes how to set up a VS Code workspace for developing VBA projects that use the ExcelSteps add-in. The setup involves managing multiple Excel files across different folders, with separate PowerShell terminals for each file to enable real-time code synchronization via xlwings.

## Project Structure

A typical project has the folder organization shown below where are developing the project_name.xlsm workbook and code. We use xlwings vba edit with project_name.xlsm, tests_project_name.xlsm and XLSteps.xlam to sync *.cls and *.bas code files with the modules in the workbooks. Although we primarily modify the project and its tests *xlsm files, It is possible that we need to edit the Excel Steps *.xlam while working. 

A special case is working on the XLSteps.xlam add-in itself with no client project involved (e.g Excel Steps is the project in this case). The Excel_Steps folder contains the src and tests subfolders and there are only two files involved (XLSteps.xlam and tests_XLSteps.xlsm) along with their synced code files.
```
Client_Projects/client_name/Development/
	project_name/                    # Master client project folder
		.github/                      # Project-specific GitHub settings/skills
		src/                          # Project source code
			project_name.xlsm          # Main project workbook
		    *.cls and *.bas files		# code files (synced with *.xlsm modules)
		tests/                        # Test suite folder for project
			tests_project_name.xlsm    # Test workbook
		    *.cls and *.bas files		# code files (synced with *.xlsm module)
		project_name.code-workspace   # VS Code workspace file
Projects/                           # Master open-sourc project folder
	Excel_Steps/                     # ExcelSteps add-in (shared across projects)
		.github/                      # ExcelSteps GitHub settings and skills
		src/
		  XLSteps.xlam                # ExcelSteps add-in file
		  *.cls and *.bas files		# code files (synced with *.xlam modules)
		tests/                        # Test suite folder for Excel_Steps add-in
		  tests_XLSteps.xlsm          # ExcelSteps Test workbook
		  *.cls and *.bas files		# code files (synced with *.xlsm module)
```

## Setup

AI should assume the starting point that the parent `project_name` folder is pre-opened by the user or by "Open Workspace From File". AI should therefore assume that .cwd is preset to be the `project_name` folder (or to Excel_Steps for special case of working on the add-in)

AI should provide minimal, non-verbose chat feedback just reporting on steps taken ("Pre-deleted Terminals named XYZ" "Successfully created Terminals XYZ"). Do not provide a verbose summary of what was done.

### 1. Create Workspace File if Non-Existent
If a code-workspace file doesn't yet exist, create `project_name.code-workspace` based on this template:
```
{
	"folders": [
		{
			"path": "."
		},
		{
			"path": "../../../../Projects/Excel_Steps"
		}
	],
	"settings": {
		"terminal.integrated.cwd": "."
	}
}
```

### 2. Close All Existing Terminals

Use Terminal: killAll VS Code instruction to close all existing terminals to ensure a clean setup state.

### 3. Create Terminals/Initialize xlwings VBA Editing for Project Files

If working on a project plus XLSteps, create ALL 3 independent background terminals sequentially with these commands:

1. Terminal 1: `cd src; xlwings vba edit -f "project_name.xlsm"`
2. Terminal 2: `cd tests; xlwings vba edit -f "tests_project_name.xlsm"`
3. Terminal 3: `cd ../../../../Projects/Excel_Steps/src; xlwings vba edit -f "XLSteps.xlam"`

If working on the XLSteps project itself, create 2 independent background terminals sequentially with these commands:

1. Terminal 1: `cd src; xlwings vba edit -f "XLSteps.xlam"`
2. Terminal 2: `cd tests; xlwings vba edit -f "tests_XLSteps.xlsm"`

After AI completes this, AI should report brief confirmation or report any errors that occurred. User will manually respond to each terminal's xlwings prompt ("Proceed? [Y/n]") with Y, then rename the terminals to "project", "tests_project", "XLSteps" and/or "tests_XLSteps"

