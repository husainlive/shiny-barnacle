# Fix Summary: Blank Rows and Summary Sheet Data Population

## Issue Description

The VBA script had two related issues affecting the output formatting and data completeness:

1. **Missing Blank Rows**: The output did not have blank rows between different profit centers, making it difficult to distinguish between profit center groups in both GL sheets and the Summary sheet.

2. **Missing Summary Sheet Data**: The Summary sheet was creating the structure (headers and row labels) but not populating the month columns with actual data values. This meant that the Summary sheet was not showing the Posted, Reversed, and Balance amounts for each month.

### Example of Incorrect Output (Before Fix):

**GL Sheet:**
```
Profit Center | Type     | Aug-24  | Dec-24 | Jan-25
10120008      | Posted   | 1000    |        | 2000
10120008      | Reversed |         | -500   |
10120008      | Balance  | 1000    | -500   | 2000
10120003      | Posted   |         | 1500   |
10120003      | Reversed |         |        |
10120003      | Balance  |         | 1500   |
```

**Summary Sheet:**
```
GL Account | Profit Center | Type     | Aug-24 | Dec-24 | Jan-25
Provisions | 10120008      | Posted   |        |        |
Provisions | 10120008      | Reversed |        |        |
Provisions | 10120008      | Balance  |        |        |
Provisions | 10120003      | Posted   |        |        |
Provisions | 10120003      | Reversed |        |        |
Provisions | 10120003      | Balance  |        |        |
```

### Example of Expected Output (After Fix):

**GL Sheet:**
```
Profit Center | Type     | Aug-24  | Dec-24 | Jan-25
10120008      | Posted   | 1000    |        | 2000
10120008      | Reversed |         | -500   |
10120008      | Balance  | 1000    | -500   | 2000

10120003      | Posted   |         | 1500   |
10120003      | Reversed |         |        |
10120003      | Balance  |         | 1500   |
```
(Note the blank row between profit centers)

**Summary Sheet:**
```
GL Account | Profit Center | Type     | Aug-24 | Dec-24 | Jan-25
Provisions | 10120008      | Posted   | 1000   |        | 2000
Provisions | 10120008      | Reversed |        | -500   |
Provisions | 10120008      | Balance  | 1000   | -500   | 2000

Provisions | 10120003      | Posted   |        | 1500   |
Provisions | 10120003      | Reversed |        |        |
Provisions | 10120003      | Balance  |        | 1500   |
```
(Note both the blank row and the populated data)

## Root Cause

### Issue 1: Missing Blank Rows

**GL Sheets:**
The code was creating profit centers sequentially without adding blank rows between them. When finding the next available row with:
```vba
pcRowPosted = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row + 1
```
It would place the new profit center immediately after the previous one.

**Summary Sheet:**
The code was incrementing `rowOut` by 1 three times (for Posted, Reversed, Balance) without adding an additional blank row afterward.

### Issue 2: Missing Summary Sheet Data

The Summary sheet creation code (lines 234-254 in original) was only setting up the structure:
```vba
wsSummary.Cells(rowOut, 1).Value = tmpGLDesc
wsSummary.Cells(rowOut, 2).Value = tmpPC
wsSummary.Cells(rowOut, 3).Value = "Posted"
```

But it was not iterating through the months to populate the actual data values from `dictData(key)(month)`.

## Solution

### Fix 1: Add Blank Row Before New Profit Centers in GL Sheets

Modified the profit center creation logic to add a blank row before each new profit center (except the first one):

```vba
If pcRowPosted = 0 Then
    pcRowPosted = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row + 1
    If pcRowPosted < 2 Then pcRowPosted = 2
    ' Add blank row before new profit center (except for the first one)
    If pcRowPosted > 2 Then pcRowPosted = pcRowPosted + 1
    wsGL.Cells(pcRowPosted, 1).Value = tmpPC
End If
```

The `If pcRowPosted > 2` check ensures we don't add a blank row before the very first profit center (row 2).

### Fix 2: Add Blank Row After Each Profit Center in Summary Sheet

Added an extra increment to `rowOut` after processing each profit center:

```vba
rowOut = rowOut + 3  ' Move past the 3 rows we just filled
' Add blank row after each profit center group
rowOut = rowOut + 1
```

### Fix 3: Populate Summary Sheet Month Data

Added a nested loop to populate the month columns with actual data:

```vba
' Fill in the month data
colNum = 4
For Each month In dictMonthsGlobal.Keys
    If dictData(key).exists(month) Then
        If dictData(key)(month) > 0 Then
            wsSummary.Cells(rowPosted, colNum).Value = Nz(wsSummary.Cells(rowPosted, colNum).Value) + dictData(key)(month)
        Else
            wsSummary.Cells(rowReversed, colNum).Value = Nz(wsSummary.Cells(rowReversed, colNum).Value) + dictData(key)(month)
        End If
        wsSummary.Cells(rowBalance, colNum).Value = Nz(wsSummary.Cells(rowPosted, colNum).Value) + Nz(wsSummary.Cells(rowReversed, colNum).Value)
    End If
    colNum = colNum + 1
Next month
```

This logic:
1. Iterates through all months in `dictMonthsGlobal`
2. For each month, checks if the current key has data for that month
3. If positive, adds to the Posted row
4. If negative, adds to the Reversed row
5. Calculates Balance as Posted + Reversed

## Why This Works

1. **Visual Separation**: Blank rows between profit centers make the output easier to read and help users quickly identify where one profit center ends and another begins.

2. **Complete Data**: The Summary sheet now displays all the data that was previously only visible in the individual GL sheets, making it a true summary of all transactions.

3. **Consistent Format**: Both the GL sheets and Summary sheet now follow the same format with blank rows, providing a consistent user experience.

4. **Reusable Nz Function**: The existing `Nz` (Null to Zero) helper function is reused to handle empty cells when calculating balances.

## Files Modified

1. **provision.vba** (lines 164-180, 234-272):
   - Added blank row logic for GL sheets (line 168)
   - Added blank row logic for Summary sheet (line 271)
   - Added month data population for Summary sheet (lines 255-267)

## Impact

- **Data Completeness**: The Summary sheet now displays all transaction data, not just the structure
- **Readability**: Blank rows between profit centers improve visual clarity in both GL sheets and Summary sheet
- **User Experience**: Users can now use the Summary sheet to get a complete overview of all transactions without having to check individual GL sheets
- **Consistency**: Output format is now consistent between GL sheets and Summary sheet

## Verification Steps

To verify the fixes work correctly:

1. Open an Excel file with the test data (multiple profit centers, multiple months, mix of positive and negative amounts)
2. Run the `BuildProvisionReports` macro
3. Check each GL sheet to verify:
   - Blank rows appear between different profit centers
   - First profit center starts at row 2 (no blank row before it)
4. Check the Summary sheet to verify:
   - All month columns are populated with actual data values
   - Posted rows show positive amounts
   - Reversed rows show negative amounts
   - Balance rows show Posted + Reversed
   - Blank rows appear between different GL/PC combinations
5. Compare values in Summary sheet with individual GL sheets to ensure data consistency

## Related Issues

This fix addresses the issues mentioned in the problem statement:
- "it seems not all the numbers getting pick" - Fixed by populating Summary sheet month data
- "also leave blank raw inbetween profit cenders" - Fixed by adding blank rows in both GL sheets and Summary sheet
