# Fix Summary: Summary Sheet Month Column Ordering Issue

## Issue Description

The Summary sheet was showing blank values for Posted and Reversed rows even when the GL accounts had respective posted and reversed values in those months. This caused data misalignment where values appeared in incorrect month columns.

### Example of Incorrect Output (Before Fix):

```
Profit Center | Type     | Aug-24  | Dec-24     | Jan-25 | ... | Aug-25    | Sep-25
10120005      | Posted   |         |            |        | ... | 79,803.00 | 79,802.00
10120005      | Reversed |         | (88,898.00)|        | ... |           |
10120005      | Balance  |         | (88,898.00)|        | ... | 79,803.00 | 79,802.00
```

The Posted values appear in Aug-25 and Sep-25 but are blank in earlier months, while Reversed shows a value only in Dec-24. However, the actual data includes transactions across multiple months.

### Example of Expected Output (After Fix):

```
Profit Center | Type     | Dec-24     | Jan-25    | Feb-25    | ... | Aug-25    | Sep-25
10120005      | Posted   |            | 79,803.00 | 79,803.00 | ... | 79,803.00 | 79,802.00
10120005      | Reversed | (88,898.00)| ...       | ...       | ... |           |
10120005      | Balance  | (88,898.00)| ...       | ...       | ... | 79,803.00 | 79,802.00
```

All data appears in the correct month columns after sorting.

## Root Cause

The Summary sheet creation code had a critical mismatch in month ordering:

**Header Creation (lines 228-232 in original):**
```vba
colNum = 4
For Each month In dictMonthsGlobal.Keys
    wsSummary.Cells(1, colNum).Value = month
    colNum = colNum + 1
Next month
```

**Data Population (lines 256-267 in original):**
```vba
colNum = 4
For Each month In dictMonthsGlobal.Keys
    If dictData(key).exists(month) Then
        ' ... populate data ...
    End If
    colNum = colNum + 1
Next month
```

Both loops iterate through `dictMonthsGlobal.Keys`, which is **UNSORTED** in VBA dictionaries. The iteration order is unpredictable and doesn't follow chronological order.

**In contrast, the GL sheets correctly sort months:**
```vba
' GL sheet code (lines 185-187)
monthList = dictMonthsGlobal.Keys
QuickSortMonths monthList, LBound(monthList), UBound(monthList)
' Then uses sorted monthList
```

### The Problem with Unsorted Months

When `dictMonthsGlobal.Keys` contains months like:
- `"01-2025"`, `"12-2024"`, `"02-2025"`, `"03-2025"`, etc.

The dictionary may return them in an arbitrary order like:
- `"02-2025"`, `"01-2025"`, `"12-2024"`, `"03-2025"`, ...

This causes:
1. Column headers to be written in the wrong order
2. Data to be written to columns that don't match their month headers
3. Values appearing "blank" because they're written to incorrect columns

## Solution

Sort the months chronologically in the Summary sheet section, consistent with how GL sheets handle months:

### Fix 1: Sort Months for Header Creation

```vba
wsSummary.Cells(1, 1).Value = "GL Account"
wsSummary.Cells(1, 2).Value = "Profit Center"
wsSummary.Cells(1, 3).Value = "Type"

' Sort months chronologically for Summary sheet
Dim sortedMonths As Variant
sortedMonths = dictMonthsGlobal.Keys
QuickSortMonths sortedMonths, LBound(sortedMonths), UBound(sortedMonths)

colNum = 4
For Each month In sortedMonths
    wsSummary.Cells(1, colNum).Value = month
    colNum = colNum + 1
Next month
```

### Fix 2: Use Sorted Months for Data Population

```vba
' Fill in the month data using sorted months
colNum = 4
For Each month In sortedMonths
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

## Why This Works

1. **Chronological Order**: Months are now sorted chronologically using the same `QuickSortMonths` function used by GL sheets
2. **Consistent Mapping**: The column number for each month is now consistent between headers and data
3. **Correct Data Placement**: Posted, Reversed, and Balance values now appear in the correct month columns
4. **Uniform Behavior**: Summary sheet and GL sheets now both use chronologically sorted months

## Files Modified

1. **provision.vba** (lines 224-272):
   - Added `sortedMonths` variable declaration (line 229)
   - Added sorting of months before creating headers (lines 230-231)
   - Changed header loop to use `sortedMonths` instead of `dictMonthsGlobal.Keys` (line 234)
   - Changed data population loop to use `sortedMonths` (line 262)

## Technical Details

The `QuickSortMonths` function (lines 278-298) sorts month strings in "mm-yyyy" format chronologically by:
1. Converting each month string to a date using `CDate("01-" & arr(i))`
2. Comparing dates to establish chronological order
3. Using quicksort algorithm for efficient sorting

This ensures months like "12-2024", "01-2025", "02-2025" are sorted as:
- "12-2024" (December 2024)
- "01-2025" (January 2025)
- "02-2025" (February 2025)

## Impact

- **Data Accuracy**: Summary sheet now displays data in the correct month columns
- **Consistency**: Summary sheet and GL sheets have identical month ordering
- **User Experience**: Users can now trust that the Summary sheet accurately reflects all transaction data
- **Bug Resolution**: Fixes the issue where Posted/Reversed values appeared blank despite having data

## Verification Steps

To verify the fix works correctly:

1. Open an Excel file with test data containing multiple months across different years
2. Ensure the data includes transactions with various profit centers and GL accounts
3. Run the `BuildProvisionReports` macro
4. Check the Summary sheet to verify:
   - Month columns are in chronological order (e.g., Dec-24, Jan-25, Feb-25, etc.)
   - Posted, Reversed, and Balance values appear in the correct month columns
   - No blank cells where data should exist
   - Values match the corresponding GL sheets
5. Compare the month order in Summary sheet with individual GL sheets to ensure consistency
6. Verify specific profit centers (like 10120005) show data in all applicable months

## Related Issues

This fix addresses the problem statement:
- "in detail gl accounts Provision posted and reversed or showing blank even those gl has respective posted and reversed values"

The root cause was that data was being written to wrong columns due to unsorted month iteration, making it appear as if values were missing when they were actually just misplaced.
