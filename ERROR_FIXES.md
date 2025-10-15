# VBA Script Error Analysis and Fixes

## Summary
This document describes the errors found in `provision.vba` and the fixes applied.

## Errors Found and Fixed

### 1. **Incorrect GL Mapping Row Assignment (Lines 100-101)**

**Problem:** 
```vba
wsMapping.Cells(wsMapping.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = GLCode
wsMapping.Cells(wsMapping.Rows.Count, 2).End(xlUp).Offset(1, 0).Value = newGLDesc
```

The code calls `wsMapping.Rows.Count` twice separately. If column 1 and column 2 have different amounts of data, this could result in the GL Code and Description being written to different rows, causing data misalignment.

**Fix:**
```vba
Dim newRow As Long
newRow = wsMapping.Cells(wsMapping.Rows.Count, 1).End(xlUp).Row + 1
wsMapping.Cells(newRow, 1).Value = GLCode
wsMapping.Cells(newRow, 2).Value = newGLDesc
```

Calculate the new row once and use it for both columns to ensure data is written to the same row.

---

### 2. **Performance Issue - Iterating Entire Column (Lines 151-156)**

**Problem:**
```vba
Dim foundRow As Range
For Each foundRow In wsGL.Range("A:A")
    If foundRow.Value = tmpPC Then
        pcRowPosted = foundRow.Row
        Exit For
    End If
Next foundRow
```

The code iterates through the entire column A (`Range("A:A")`), which includes all 1,048,576 rows in Excel. This is extremely inefficient and can cause significant performance issues.

**Fix:**
```vba
Dim foundRow As Long
Dim lastUsedRow As Long
lastUsedRow = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row
If lastUsedRow < 1 Then lastUsedRow = 1
For foundRow = 2 To lastUsedRow
    If wsGL.Cells(foundRow, 1).Value = tmpPC Then
        pcRowPosted = foundRow
        Exit For
    End If
Next foundRow
```

Only iterate through the used range of column A, significantly improving performance.

---

### 3. **Missing GL Account Values in Summary Sheet (Lines 234, 238)**

**Problem:**
```vba
wsSummary.Cells(rowOut, 1).Value = tmpGLDesc
wsSummary.Cells(rowOut, 2).Value = tmpPC
wsSummary.Cells(rowOut, 3).Value = "Posted"
rowOut = rowOut + 1
wsSummary.Cells(rowOut, 2).Value = tmpPC
wsSummary.Cells(rowOut, 3).Value = "Reversed"
rowOut = rowOut + 1
wsSummary.Cells(rowOut, 2).Value = tmpPC
wsSummary.Cells(rowOut, 3).Value = "Balance"
```

The "Reversed" and "Balance" rows were missing the GL Account value in column 1, creating incomplete data in the Summary sheet.

**Fix:**
```vba
wsSummary.Cells(rowOut, 1).Value = tmpGLDesc
wsSummary.Cells(rowOut, 2).Value = tmpPC
wsSummary.Cells(rowOut, 3).Value = "Posted"
rowOut = rowOut + 1
wsSummary.Cells(rowOut, 1).Value = tmpGLDesc
wsSummary.Cells(rowOut, 2).Value = tmpPC
wsSummary.Cells(rowOut, 3).Value = "Reversed"
rowOut = rowOut + 1
wsSummary.Cells(rowOut, 1).Value = tmpGLDesc
wsSummary.Cells(rowOut, 2).Value = tmpPC
wsSummary.Cells(rowOut, 3).Value = "Balance"
```

Added the GL Account value for all three rows to maintain data consistency.

---

### 4. **Missing Type Labels for Existing Profit Centers (Lines 166-168)**

**Problem:**
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

The code only sets the Type labels ("Posted", "Reversed", "Balance") when creating a NEW Profit Center block. When an existing Profit Center is found, these labels are not refreshed or verified, which can lead to incomplete output where only the "Posted" row appears, and the "Reversed" and "Balance" rows are missing their labels.

**Fix:**
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
```

Move the Type label assignment outside the `If pcRowPosted = 0` block to ensure that all three rows (Posted, Reversed, Balance) are properly labeled for both new and existing Profit Centers. This ensures the complete 3-row structure is always present in the output.

---

### 5. **Worksheet Object Not Reset Before Access with Error Handling (Lines 21, 142, 208)**

**Problem:**
```vba
On Error Resume Next
Set wsGL = wb.Sheets(tmpGLDesc)
If wsGL Is Nothing Then
    Set wsGL = wb.Sheets.Add
    wsGL.Name = tmpGLDesc
```

When `On Error Resume Next` is active and you attempt to access a non-existent sheet with `Set wsGL = wb.Sheets(tmpGLDesc)`, VBA does not automatically set `wsGL` to `Nothing` on error. Instead, the object variable retains its previous value. This causes a critical bug where:
- The first GL sheet is created successfully
- When processing a second GL account, `wsGL` still points to the first sheet
- The `If wsGL Is Nothing` check fails (because it's not Nothing)
- Data for all GL accounts gets written to the same first sheet
- No additional sheets are created

**Fix:**
```vba
On Error Resume Next
Set wsGL = Nothing
Set wsGL = wb.Sheets(tmpGLDesc)
If wsGL Is Nothing Then
    Set wsGL = wb.Sheets.Add
    wsGL.Name = tmpGLDesc
```

Explicitly set the worksheet object to `Nothing` before attempting to access it. This ensures that if the sheet doesn't exist, the object variable will remain `Nothing`, allowing the subsequent check to work correctly and create new sheets as needed.

This fix was applied to three locations:
- Line 21: `wsMapping` for GL_Mapping sheet
- Line 142: `wsGL` for individual GL account sheets (main bug fix)
- Line 208: `wsSummary` for Summary sheet

---

### 6. **Missing Profit Center Values in Reversed and Balance Rows (Lines 169-172)**

**Problem:**
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
```

The code only sets the Profit Center value (column A) for the Posted row when creating a new Profit Center block. The Reversed and Balance rows (pcRowPosted + 1 and pcRowPosted + 2) do not have the Profit Center value set in column A, resulting in output where these rows appear with blank Profit Center cells. This creates a format like:

```
Profit Center  Type      Aug-24  Sep-25
10120007       Posted    1000    1000
               Reversed  -500    -500
               Balance   500     500
```

Instead of the expected format:

```
Profit Center  Type      Aug-24  Sep-25
10120007       Posted    1000    1000
10120007       Reversed  -500    -500
10120007       Balance   500     500
```

**Fix:**
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
```

Set the Profit Center value in column A for all three rows (Posted, Reversed, Balance), not just the Posted row. This ensures that all three rows display the Profit Center value, making the output consistent and properly formatted.

---

## Impact

1. **Data Integrity:** Fix #1, #3, #4, #5, and #6 ensure data is written correctly and completely
2. **Performance:** Fix #2 can reduce execution time from minutes to seconds for large datasets
3. **Reliability:** All fixes improve the overall reliability and correctness of the script
4. **Output Completeness:** Fix #4 ensures that all three rows (Posted, Reversed, Balance) are properly displayed for each Profit Center, matching the expected output format
5. **Multi-Sheet Support:** Fix #5 ensures that multiple GL account sheets are created correctly instead of overwriting data into a single sheet
6. **Data Consistency:** Fix #6 ensures that all three rows (Posted, Reversed, Balance) display the Profit Center value in column A, preventing blank cells and maintaining proper data structure

## Testing Recommendations

1. Test with a GL_Mapping sheet containing different amounts of data in columns 1 and 2
2. Test with worksheets containing large amounts of data (1000+ rows)
3. Verify Summary sheet output contains complete GL Account information for all row types
4. Verify GL sheets are created correctly with proper Profit Center blocks
5. **Test with multiple GL accounts to ensure separate sheets are created for each GL account**
6. **Verify that data is not overwritten or combined into a single sheet when multiple GL accounts exist**

## Notes

- The script assumes that the GL_Mapping sheet in Personal.xlsb exists and has proper structure (headers in row 1)
- The script creates headers for new GL sheets automatically, so they will never be completely empty
- All fixes maintain the existing assumptions and error handling patterns of the original script
