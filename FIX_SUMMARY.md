# Fix Summary: Individual Expenses Output Issues

## Issue Description
The individual expenses output had two related issues:

**Issue 1 (Fix #4):** The output was displaying only "Posted" rows for each Profit Center with all amounts combined, instead of showing the proper 3-row structure:
- Posted (positive amounts)
- Reversed (negative amounts)
- Balance (sum of Posted + Reversed)

**Issue 2 (Fix #6):** The Reversed and Balance rows were missing Profit Center values in column A, showing blank cells instead of the Profit Center number.

### Example of Incorrect Output (Before Fixes):
```
Profit Center | Type   | Aug-24    | Oct-24    | Nov-24    | Dec-24
10120002      | Posted |           |           | 41200     |
10120007      | Posted |           | -1770.3   | -2150     | -42399.97
```
Or when some rows existed but were incomplete:
```
Profit Center | Type     | Aug-24 | Oct-24 | Nov-24 | Dec-24
10120002      | Posted   | 1000   | 1000   | 1000   | 1000
              | Reversed | -1000  | 0      | -1000  | 0
              | Balance  | 0      | 1000   | 0      | 1000
```

### Example of Expected Output (After Fixes):
```
Profit Center | Type     | Aug-24 | Oct-24 | Nov-24 | Dec-24
10120002      | Posted   | 1000   | 1000   | 1000   | 1000
10120002      | Reversed | -1000  | 0      | -1000  | 0
10120002      | Balance  | 0      | 1000   | 0      | 1000
```

## Root Cause

### Fix #4: Missing Type Labels
In the original code (lines 162-169), the Type labels ("Posted", "Reversed", "Balance") were only set when creating a NEW Profit Center block:

```vba
If pcRowPosted = 0 Then
    pcRowPosted = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row + 1
    If pcRowPosted < 2 Then pcRowPosted = 2
    wsGL.Cells(pcRowPosted, 1).Value = tmpPC
    wsGL.Cells(pcRowPosted, 2).Value = "Posted"
    wsGL.Cells(pcRowPosted + 1, 2).Value = "Reversed"
    wsGL.Cells(pcRowPosted + 2, 2).Value = "Balance"
End If
```

When an existing Profit Center was found (lines 156-161), the code would:
1. Find the existing row number
2. Skip the Type label assignment
3. Proceed to fill data into rows that might not have proper labels

This caused the Reversed and Balance rows to appear empty or unlabeled in the output.

### Fix #6: Missing Profit Center Values
The code only set the Profit Center value (column A) for the Posted row when creating a new Profit Center block. The Reversed and Balance rows did not have their Profit Center values set, resulting in blank cells in column A for those rows.

## Solution

### Fix #4: Move Type Label Assignment Outside Conditional Block
Move the Type label assignment outside the `If pcRowPosted = 0` block to ensure labels are set for both new AND existing Profit Centers:

```vba
If pcRowPosted = 0 Then
    pcRowPosted = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row + 1
    If pcRowPosted < 2 Then pcRowPosted = 2
    wsGL.Cells(pcRowPosted, 1).Value = tmpPC
End If
' Always ensure the Type labels are set for all 3 rows
wsGL.Cells(pcRowPosted, 2).Value = "Posted"
wsGL.Cells(pcRowPosted + 1, 2).Value = "Reversed"
wsGL.Cells(pcRowPosted + 2, 2).Value = "Balance"
pcRowReversed = pcRowPosted + 1
pcRowBalance = pcRowPosted + 2
```

### Fix #6: Set Profit Center Value for All 3 Rows
Add assignments to set the Profit Center value in column A for all three rows (Posted, Reversed, Balance):

```vba
If pcRowPosted = 0 Then
    pcRowPosted = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row + 1
    If pcRowPosted < 2 Then pcRowPosted = 2
    wsGL.Cells(pcRowPosted, 1).Value = tmpPC
End If
' Always ensure the Profit Center value is set for all 3 rows
wsGL.Cells(pcRowPosted, 1).Value = tmpPC
wsGL.Cells(pcRowPosted + 1, 1).Value = tmpPC
wsGL.Cells(pcRowPosted + 2, 1).Value = tmpPC
' Always ensure the Type labels are set for all 3 rows
wsGL.Cells(pcRowPosted, 2).Value = "Posted"
wsGL.Cells(pcRowPosted + 1, 2).Value = "Reversed"
wsGL.Cells(pcRowPosted + 2, 2).Value = "Balance"
pcRowReversed = pcRowPosted + 1
pcRowBalance = pcRowPosted + 2
```

## Why This Works
1. **New Profit Centers**: When `pcRowPosted = 0`, the code creates a new row and sets the Profit Center value, then sets both the Profit Center values and Type labels for all 3 rows
2. **Existing Profit Centers**: When `pcRowPosted` is found (not 0), the code skips row creation but still sets the Profit Center values and Type labels for all 3 rows
3. **Data Population**: The existing data population logic (lines 197-205) already correctly:
   - Adds positive amounts to the Posted row
   - Adds negative amounts to the Reversed row
   - Calculates Balance as Posted + Reversed

With both fixes applied:
- All three rows now have proper Type labels in column B
- All three rows now have Profit Center values in column A
- The output format matches expectations with complete data in all rows

## Files Modified
1. **provision.vba** (lines 162-178): 
   - Moved Type label assignment outside conditional block (Fix #4)
   - Added Profit Center value assignment for all 3 rows (Fix #6)
2. **ERROR_FIXES.md**: Added documentation as Fix #4 and Fix #6
3. **TEST_SCENARIO.md**: Created test scenario documentation

## Verification Steps
To verify the fix works correctly:

1. Open an Excel file with transaction data
2. Ensure data includes:
   - Mix of PostingKey 50 (positive) and 40 (negative) transactions
   - Multiple Profit Centers
   - Multiple months
3. Run the `BuildProvisionReports` macro
4. Check each GL sheet to verify:
   - Each Profit Center has exactly 3 rows
   - Row 1 (Posted) shows only positive amounts
   - Row 2 (Reversed) shows only negative amounts
   - Row 3 (Balance) shows Posted + Reversed for each month
   - Type column (B) shows "Posted", "Reversed", "Balance" labels
   - **Profit Center column (A) shows the Profit Center value for all 3 rows (Posted, Reversed, and Balance)**

## Impact
- **Data Completeness**: All three rows (Posted, Reversed, Balance) now appear for every Profit Center with complete Profit Center values in column A
- **Data Accuracy**: Amounts are correctly separated into positive (Posted) and negative (Reversed)
- **Reliability**: Type labels and Profit Center values are always set, even when processing existing Profit Centers
- **User Experience**: Output matches expected format described in the problem statement with no missing Profit Center values
