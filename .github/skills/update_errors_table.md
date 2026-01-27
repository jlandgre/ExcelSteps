# Update VBA Errors Table

Generate or update an errors table for VBA code modules with the following structure:
- Columns: `iCode`, `Module`, `Routine`, `Msg_String`, `sVal`, `iMsgDevUser`
- Tab-separated values (single tab between columns)

## Requirements

### For Table Updates
1. **Base lines**: Include `Msg_String = "Base"` for each VBA subroutine/function in order they appear in code
2. **Add new routines**: Create Base lines for new subroutines/functions not in current table
3. **Remove deleted**: Delete lines where Routine no longer exists in code module
4. **Preserve iCode numbering**: Start at same value as original table's first routine
5. **Increment iCode**: Use increments of 10 between Base lines (unless routine has >9 secondary lines, then increment by 20)
6. **Secondary lines**: Retain non-Base lines for same Routine, adjusting iCode to maintain offset from new Base value
   - Example: If routine "MySub" Base changes from 120→200, secondary line 121 becomes 201
7. **Module column**: Set to provided module name
8. **iMsgDevUser column**: 
   - FALSE for all Base lines
   - TRUE for secondary lines (unless explicitly set to FALSE in original)
9. **sVal column**: Leave empty for all rows
10. **Format**: Single tab between columns; multiple tabs only for blank column values

### For New Tables (From Scratch)
- Create Base lines only (one per sub/function)
- Use specified starting iCode (e.g., 2500)
- Increment by 10s between routines
- Set iMsgDevUser = FALSE for all Base lines
- Set Module to specified name

## Usage Examples

**Update existing table:**
```
Update the errors table for module "modValidation" using the code below:
<provide table and code>
```

**Create new table:**
```
Build errors table from scratch for module "modUtilities", starting at iCode 3000:
<provide code>
```

## Output
- Respond "No changes needed" if table is already correct
- Otherwise, output complete revised tab-separated table
