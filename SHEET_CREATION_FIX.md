# Fix Summary: Multiple GL Sheet Creation Issue

## Issue Description
The VBA script was creating only ONE sheet instead of creating separate sheets for each GL account. All data from multiple GL accounts was being written to the same first sheet, causing data to be overwritten and mixed together.

### Example of Incorrect Behavior (Before Fix):
When processing data with 3 different GL accounts (e.g., "Office Supplies", "Equipment", "Services"):
- Only "Office Supplies" sheet is created
- Data for "Equipment" and "Services" is written to the "Office Supplies" sheet
- The "Equipment" and "Services" sheets are never created
- All profit centers from all GL accounts are mixed together in one sheet

### Example of Expected Behavior (After Fix):
When processing data with 3 different GL accounts:
- "Office Supplies" sheet is created with its profit centers
- "Equipment" sheet is created with its profit centers  
- "Services" sheet is created with its profit centers
- Each sheet contains only the profit centers belonging to that GL account

## Root Cause
In VBA, when `On Error Resume Next` is active and you attempt to access a non-existent worksheet using:
```vba
Set wsGL = wb.Sheets(tmpGLDesc)
```

If the sheet doesn't exist, VBA does **not** automatically set `wsGL` to `Nothing`. Instead, the object variable **retains its previous value**.

This caused the following sequence of events in the original code (lines 140-149):

1. **First iteration** (GL account: "Office Supplies"):
   - `Set wsGL = wb.Sheets("Office Supplies")` → fails, but `wsGL` is still `Nothing` (initial state)
   - `If wsGL Is Nothing` → TRUE
   - Creates new sheet "Office Supplies"
   - `wsGL` now points to "Office Supplies" sheet

2. **Second iteration** (GL account: "Equipment"):
   - `Set wsGL = wb.Sheets("Equipment")` → fails, but `wsGL` **still points to "Office Supplies"**
   - `If wsGL Is Nothing` → FALSE (it's not Nothing!)
   - Code thinks sheet exists, skips creation
   - Data for "Equipment" is written to "Office Supplies" sheet

3. **Third iteration** (GL account: "Services"):
   - Same problem: `wsGL` still points to "Office Supplies"
   - Data for "Services" is also written to "Office Supplies" sheet

## Solution
Explicitly reset the worksheet object to `Nothing` before attempting to access it:

```vba
' Create or activate GL sheet
On Error Resume Next
Set wsGL = Nothing          ' ← ADDED: Reset to Nothing
Set wsGL = wb.Sheets(tmpGLDesc)
If wsGL Is Nothing Then
    Set wsGL = wb.Sheets.Add
    wsGL.Name = tmpGLDesc
    wsGL.Range("A1").Value = "Profit Center"
    wsGL.Range("B1").Value = "Type"
End If
On Error GoTo 0
```

By adding `Set wsGL = Nothing` before accessing the sheet, we ensure that:
- If the sheet exists, `wsGL` is set to that sheet
- If the sheet doesn't exist, `wsGL` remains `Nothing`
- The `If wsGL Is Nothing` check works correctly in both cases

## Why This Works
1. **First iteration** (GL account: "Office Supplies"):
   - `Set wsGL = Nothing` → clears the variable
   - `Set wsGL = wb.Sheets("Office Supplies")` → fails, `wsGL` stays `Nothing`
   - `If wsGL Is Nothing` → TRUE
   - Creates "Office Supplies" sheet

2. **Second iteration** (GL account: "Equipment"):
   - `Set wsGL = Nothing` → **clears reference to "Office Supplies" sheet**
   - `Set wsGL = wb.Sheets("Equipment")` → fails, `wsGL` stays `Nothing`
   - `If wsGL Is Nothing` → TRUE
   - Creates "Equipment" sheet correctly

3. **Third iteration** (GL account: "Services"):
   - `Set wsGL = Nothing` → clears reference
   - Creates "Services" sheet correctly

## Files Modified
1. **provision.vba** (line 142): Added `Set wsGL = Nothing` for GL account sheet creation
2. **provision.vba** (line 21): Added `Set wsMapping = Nothing` for GL_Mapping sheet access (preventive fix)
3. **provision.vba** (line 208): Added `Set wsSummary = Nothing` for Summary sheet access (preventive fix)
4. **ERROR_FIXES.md**: Added documentation as Fix #5
5. **SHEET_CREATION_FIX.md**: Created this detailed documentation

## Verification Steps
To verify the fix works correctly:

1. Prepare test data with **multiple GL accounts** (at least 3 different GL codes)
2. Each GL account should have multiple transactions across different profit centers and months
3. Run the `BuildProvisionReports` macro
4. Verify that:
   - **Multiple sheets are created** (one per GL account)
   - Each sheet is named after its GL account description
   - Each sheet contains only the profit centers belonging to that GL account
   - Data is not mixed between different GL accounts
   - The Summary sheet lists all GL accounts correctly

## Impact
- **Critical Bug Fix**: Resolves the main issue reported in the problem statement
- **Data Integrity**: Ensures data from different GL accounts is properly separated
- **Correct Output**: Multiple GL sheets are now created as intended by the script design
- **User Experience**: Users can now properly analyze data per GL account in separate sheets
- **Preventive**: Applied the same fix to other worksheet access points to prevent similar issues

## Related Issues
This fix directly addresses the problem statement:
> "when i execte it seems value over written it creates only one sheet even there are many expenses existis in the sheet"

The root cause was that object variables were not being reset, causing all expenses to be written to a single sheet regardless of how many different GL accounts existed in the data.
