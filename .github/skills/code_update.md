# Skill: Code Update Process

## Overview
This skill documents the process for making updates to Dashboard VBA project code, from initial planning through implementation.

## Workflow Steps

### 1. Initial Planning Mode

**User Actions:**
- Create new file: `docs/Change_Log_[description]_[MMDDYY].md`
- Paste description of changes into the file under `## Description of Change` section
- Notify agent to begin analysis

**Agent Actions:**
- Read the change log description
- Analyze relevant project and test code modules
- Identify affected functions, classes, and modules
- Analyze tblRowsCols and mdlScenario provisioned status when making changes (e.g. can't access a tbl.rngRows attribute if the tbl has not been previously provisioned)
- **Ask clarifying questions** before proceeding (DO NOT assume or guess)

### 2. Question & Clarification Phase

**Agent Actions:**
- Present findings from code analysis
- Ask specific questions about:
  - Ambiguous requirements
  - Missing details (e.g., new function names, default values, behavior specifics)
  - Edge cases or implications
  - Preferred implementation approach choices
- List all affected modules and functions discovered

**User Actions:**
- Answer agent's questions
- Clarify any ambiguous requirements
- Approve or modify suggested approach

### 3. Architectural Planning

**Agent Actions:**
- Add `## Architectural Approach` section to Change_Log file
- Document whether changes are:
  - Simple modifications to existing functions
  - Larger refactoring with cascading impacts
  - New functionality requiring new modules/classes
- List design decisions and rationale

### 4. Impact Documentation

**Agent Actions:**
- Add `## Affected Code` section with tables such as this example:

#### Project Code Changes
| Module | Function | Change Type | Description |
|--------|----------|-------------|-------------|
| CurSnap.cls | FunctionName | Modify/Rename/New | Brief description |

#### Test Code Changes  
| Module | Test Function | Change Type | Description |
|--------|---------------|-------------|-------------|
| Tests_CurSnap.bas | test_FunctionName | Modify/Rename | Brief description |

**Change Types:**
- **Modify**: Update existing function logic
- **Rename**: Change function name (update all call sites)
- **New**: Create new function or test
- **Delete**: Remove deprecated function

### 5. Implementation Authorization

**User Actions:**
- Review the documented approach and impact analysis
- Authorize agent to proceed: "Proceed with code changes"

### 6. Implementation Planning

**Agent Actions:**
- Add `## Implementation Todo List` section to Change_Log file
- Use `manage_todo_list` tool to create task list with items like:
  1. Update function X in module Y
  2. Rename function A to B (N call sites)
  3. Add tests for new functionality
  4. Update VersionDashboard constant
  5. Run test suite to verify

### 7. Implementation Execution

**Agent Actions:**
- Work through todo list systematically
- Mark each item in-progress, then completed
- Make code changes using multi_replace_string_in_file when possible
- **CRITICAL**: Update version date in `tests/Constants.bas`:
  ```vba
  Public Const VersionDashboard As String = "version M/DD/YY"
  ```
- Document any deviations or issues encountered

### 8. Verification & Completion

**Agent Actions:**
- Add `## Implementation Summary` section to Change_Log file
- Document:
  - All changes made
  - Any deviations from plan
  - Test results (if run)
  - Known issues or follow-ups needed
- Check for compilation errors using `get_errors` tool

**User Actions:**
- Run test suite to verify changes
- Test functionality manually if needed
- Mark Change_Log as complete or request revisions

## Best Practices

1. **Never skip the question phase** - Clarify all ambiguities before coding
2. **Update version constant** - Always update `VersionDashboard` in tests/Constants.bas
3. **Test coverage** - Ensure renamed functions have corresponding test updates
4. **Small batches** - For large refactorings, break into smaller change logs
5. **Document deviations** - If implementation differs from plan, document why
6. **Preserve patterns** - Follow project's VBA architectural patterns (see copilot-instructions.md)

## Large-Scale Code Deletion Strategies

When deleting multiple functions or large blocks of code from a file, use these strategies to avoid file corruption and line number targeting errors:

### Backward Deletion Approach (Recommended)

**Strategy**: Delete code blocks working from the END of the file toward the BEGINNING.

**Why This Works:**
- Prevents line number shifts for remaining code to be deleted
- Each deletion only affects line numbers BELOW the deletion point
- Code above the deletion point retains its original line numbers
- Minimizes risk of targeting errors in subsequent deletions

**Implementation Steps:**
1. **Catalog all deletions**: Use grep_search or subagent to identify ALL functions/blocks to delete with their line numbers
2. **Sort deletions by line number**: Organize from highest line number to lowest
3. **Delete in batches**: Work backwards through the file in batches (5-10 functions per batch is manageable)
4. **Optional verification**: For critical operations, use grep_search between batches to confirm progress

**Example Workflow:**
```
Functions to delete at lines: 165, 295, 352, 479, 592, 688, 720, 756, 784
Delete order:
  Batch 1: Line 784 (highest)
  Batch 2: Line 756
  Batch 3: Line 720
  Batch 4: Lines 688, 620 (group nearby functions)
  Batch 5: Line 592
  Batch 6: Lines 479-520 (group range)
  Batch 7: Lines 295-352 (group range)
  Batch 8: Line 165 (lowest)
```

### When to Use Step-Wise Deletion

Use step-wise backward deletion when:
- Deleting 10+ functions or blocks from a single file
- File has shown corruption issues with large deletions
- Functions are interspersed with other code to preserve
- Previous deletion attempts resulted in line number targeting errors

### Single Large Deletion (Use With Caution)

Large single deletions (100+ lines) can work IF:
- The deleted code is a contiguous block with clear boundaries
- No subsequent deletions are needed in the same operation
- Context strings are unique and unambiguous

### Verification Steps

After completing large-scale deletions:
1. **Use grep_search**: Confirm all target code removed (search for function names or unique patterns)
2. **Use get_errors**: Verify no compilation errors introduced
3. **Review file length**: Confirm expected line count reduction (e.g., 1846 lines → ~925 lines)

### Troubleshooting Corruption

If file corruption occurs despite these strategies:
- Check for duplicate function declarations (sign of failed deletions)
- Check for functions with mismatched names/bodies (another corruption indicator)
- If corruption is detected, request step-wise backward deletion approach from the user
- Start fresh from end of file working upward to clean up corruption systematically

## Common Gotchas

- **Function renaming**: Must update all call sites in project AND tests
- **Test module updates**: Renamed functions need renamed test functions
- **Version constant**: Easy to forget - add to todo list
- **Error handling**: Maintain `SetErrs`/`GoTo ErrorExit` pattern in all functions
- **Docstrings**: Update function docstrings when behavior changes
- **Large deletions**: Use backward deletion strategy when removing 10+ code blocks to prevent corruption