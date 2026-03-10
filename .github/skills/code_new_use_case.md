# Code new use case

## Purpose
Define a concise, repeatable process for planning and implementing new VBA project use cases while preserving VBA architecture and testability.

## Core guidance
Use the data object architecture in `copilot-instructions.md` (`tblRowsCols` + `mdlScenario`).

Use `code_plan.xlsx` (AI sees `.csv`) to track procedures, sub-methods, and test hooks.

Planning notes are drafted by AI and refined collaboratively; they are both pre-work and durable documentation. Project includes a code project folder and an Obsidian graph folder. Planning mode documents should be stored within Obsidan folder and use Obsidian formatting

Steps to be completed by human project owner (e.g. non-AI) are denoted by "(ProjOwner)"

## Workflow
1. (ProjOwner) Write description of use case in `UseCase_new_use_case_name.md` Obsidian note
2. Define scope and target outputs. Typically, this is based on in-depth AI review of a projname_Architecture.md Obsidian graph note.
3. Draft parent use-case note (e.g., `UseCase_new_use_case_name.md`. See sections below).
4. Draft Stage notes (one note per stage procedure. See sections below).
5. Draft sub-procedure notes only for complex steps.
6. Sync method names and sequence into `code_plan.xlsx` see [[Example_Code_Plan]] in Obsidian graph folder.
7. Implement after planning notes and code plan agree.

## Planning note templates

### 1) Parent use-case note (e.g., `UseCase_new_use_case_name.md`)
Use when describing end-to-end flow.

UseCase Note Sections:
- Purpose
- Background
- Data Shape Assumptions
- Architecture (classes)
- Transform and Mapping Plan (Stage 1..n)
- Test Plan

Stage Note formatting:
- `Stage N: Name ([[ParentStageProcedure]])`
- list sequential method names under each stage
- include child note links only for complex sub-steps


### 3) Sub-procedure Notes (e.g., `MapHeadersToCanonicalNames.md`)
Use only when a stage step is complex enough to need drill-down.

Required sections:
- Goal
- Scope
- Procedure Arguments
- Sequential Sub-Steps
- One-Line Responsibility by Sub-Step
- Output Contract
- Error/Warning Model
- Minimal Test Focus
- Related Notes

## Language and formatting standards
- Prefer transparent language: say what runs first/next and what outputs are produced.
- Be concise: avoid repeated sections and avoid long narrative paragraphs.
- Use stable method names and explicit sequence numbers.
- Keep canonical names aligned to Scenario Model variable names. ((Canonical names are variables' "names of record" which may differ from import names and variable labels/descriptions) )
- Keep Obsidian-compatible links with `[[WikiLink]]` style.
- Avoid absolute local paths and tool-specific syntax in vault notes.

## Naming and planning conventions
- Favor explicit procedure names ending in `Procedure` for top-level workflows.
- Use camel case method names that describe one responsibility.
- Keep canonical field names aligned to Scenario Model variable names.
- Keep planning notes and code plan synchronized before coding.

## Related notes
- [[Example_Code_Plan]].md in Obsidian graph folder
