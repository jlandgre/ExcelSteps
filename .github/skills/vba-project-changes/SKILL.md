---
name: vba-project-changes
description: Plan and implement VBA project changes using a staged Change_ note workflow aligned to Dashboard and ExcelSteps architecture. Use when user requests a new VBA feature, refactor, or test expansion and wants planning through implementation.
---

# VBA Project Changes

## Purpose
Define a concise, repeatable process for planning and implementing code changes in a VBA project and its test suite.

## Core Guidance
- Refer to collaborators as `ProjectOwner` (human) and `AI`.
- Follow project data-object architecture in `.github/copilot-instructions.md` (`tblRowsCols` and `mdlScenario`).
- Workspace context typically includes:
  - `ProjectName` folder (code)
  - `Graph_ProjectName` Obsidian folder (design/change notes)
  - Optional `ExcelSteps` folder if used by project code
- Skills and `copilot-instructions.md` are sourced from `ProjectName/.github`.

## Stage-Gated Workflow
1. Discovery
- Read the request and relevant code/docs.
- Read the current `Change_ChangeName.md` if it exists.
- Follow linked background notes before drafting implementation details.

2. Planning (required)
- Use `grill-me` style questioning to resolve design branches and dependencies.
- Draft `Change_ChangeName.md` from answers.
- Stop at this gate until `ProjectOwner` explicitly approves the `Change_` note.

3. Implementation Design
- Produce/update procedure outlines and supporting `procPlan_*.md` notes.
- Align architecture decisions with `VBA Project Architecture` expectations.
- Flag legacy non-conforming patterns and request scope decision on refactor vs leave-as-is.

4. Code + Tests
- Implement project code changes and matching tests.
- Keep test architecture explicit: module location, Procedures grouping, and coverage scope.
- Validate cross-workbook factory instantiation needs for new classes.

## Required Sections for `Change_ChangeName.md`
1. Purpose: high-level change objective.
2. Background: linked architecture/context docs and change-specific framing.
3. Data I/O Descriptions: source/target data structures, mappings, and key arguments.
4. Project Architecture: new/modified classes and responsibilities.
5. Test Architecture: test module location and Procedure groups in `tests_ProjName.xlsm`.
6. Discussion: Topic XYZ (optional): key options and design tradeoffs.
7. Testing Considerations:
- Module structure and target procedures.
- `Procedures.cls` grouping attributes.
- Unit/integration coverage expectations.
- Existing tests affected.
- Test data file requirements (`test_data_xxx` folders).
- Cross-workbook factory function requirements.
- Edge cases and boundary validations.
8. Procedure Outline: top-level procedure flow and sub-procedure/method list.

## Procedure Outline Expectations
- Include ordered flow from entry procedure through helper methods.
- Link to `procPlan_ProcedureMethodName.md` where deeper planning is needed.
- Use short draft docstring-style descriptions for each method.
- Identify reusable sub-procedures when logic should be shared across use cases.

## Completion Criteria
- `Change_` note approved by `ProjectOwner`.
- Architecture and test plan documented.
- Procedure outline complete and linked where needed.
- Code and tests implemented with architecture-consistent patterns.
