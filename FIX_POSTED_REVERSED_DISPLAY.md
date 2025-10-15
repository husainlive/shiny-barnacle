# Fix Summary: Display Actual Posted and Reversed Values

## Issue Description

When a provision was posted (positive amount) and then subsequently reversed (negative amount) in the same month, the system would show zero in both the Posted and Reversed rows instead of showing the actual posted and reversed values separately.

### Example of Incorrect Output (Before Fix):

For transactions in August:
- Transaction 1: GL=Provisions, PC=10120002, Amount=+1000 (Posted)
- Transaction 2: GL=Provisions, PC=10120002, Amount=-1000 (Reversed)

**Current Output:**
```
Profit Center | Type     | Aug-24 | Total
10120002      | Posted   |        | 0
10120002      | Reversed |        | 0
10120002      | Balance  | 0      | 0
```

The amounts were being netted to zero in the data collection phase, so neither the Posted nor Reversed values were visible.

### Example of Expected Output (After Fix):

**Expected Output:**
```
Profit Center | Type     | Aug-24  | Total
10120002      | Posted   | 1000    | 1000
10120002      | Reversed | -1000   | -1000
10120002      | Balance  | 0       | 0
```

Now both the Posted (+1000) and Reversed (-1000) values are visible, even though they net to zero in the Balance row.

## Root Cause

The issue was in the data collection logic (lines 114-119 in the original code):

```vba
' Sum amounts per month
If dictData(key).exists(tmpMonthKey) Then
    dictData(key)(tmpMonthKey) = dictData(key)(tmpMonthKey) + tmpAmount
Else
    dictData(key)(tmpMonthKey) = tmpAmount
End If
```

This logic **summed all amounts** for the same GL+Profit Center+Month combination into a single value. When both positive and negative amounts existed for the same month, they would cancel each other out:
- +1000 + (-1000) = 0

The output logic (lines 208-218) then checked if this summed value was positive or negative to determine whether to display it in the Posted or Reversed row. Since the sum was 0, nothing was displayed.

## Solution

Changed the dictionary structure to **track Posted and Reversed amounts separately** from the data collection phase:

### Change 1: Separate Tracking in Data Collection (Lines 114-126)

```vba
' Track Posted and Reversed amounts separately per month
If Not dictData(key).exists(tmpMonthKey) Then
    Set dictData(key)(tmpMonthKey) = CreateObject("Scripting.Dictionary")
    dictData(key)(tmpMonthKey)("Posted") = 0
    dictData(key)(tmpMonthKey)("Reversed") = 0
End If

' Add to Posted or Reversed based on sign
If tmpAmount > 0 Then
    dictData(key)(tmpMonthKey)("Posted") = dictData(key)(tmpMonthKey)("Posted") + tmpAmount
Else
    dictData(key)(tmpMonthKey)("Reversed") = dictData(key)(tmpMonthKey)("Reversed") + tmpAmount
End If
```

Now the dictionary structure is:
```
dictData("Provisions|10120002")("08-2024")("Posted") = 1000
dictData("Provisions|10120002")("08-2024")("Reversed") = -1000
```

Instead of:
```
dictData("Provisions|10120002")("08-2024") = 0
```

### Change 2: Read Separate Values in Output Logic (Lines 211-235)

```vba
' Fill Posted/Reversed/Balance and track totals
totalPosted = 0
totalReversed = 0

For Each month In dictData(key).Keys
    colNum = monthsDict(month)
    Dim postedAmt As Double, reversedAmt As Double
    postedAmt = dictData(key)(month)("Posted")
    reversedAmt = dictData(key)(month)("Reversed")
    
    ' Always write Posted value if non-zero
    If postedAmt <> 0 Then
        wsGL.Cells(pcRowPosted, colNum).Value = Nz(wsGL.Cells(pcRowPosted, colNum).Value) + postedAmt
    End If
    totalPosted = totalPosted + postedAmt
    
    ' Always write Reversed value if non-zero
    If reversedAmt <> 0 Then
        wsGL.Cells(pcRowReversed, colNum).Value = Nz(wsGL.Cells(pcRowReversed, colNum).Value) + reversedAmt
    End If
    totalReversed = totalReversed + reversedAmt
    
    ' Calculate Balance
    wsGL.Cells(pcRowBalance, colNum).Value = Nz(wsGL.Cells(pcRowPosted, colNum).Value) + Nz(wsGL.Cells(pcRowReversed, colNum).Value)
Next month
```

### Change 3: Update Summary Sheet Logic (Lines 322-326)

```vba
' Sum all months for this GL+PC combination
For Each month In dictData(key).Keys
    totalPostedVal = totalPostedVal + dictData(key)(month)("Posted")
    totalReversedVal = totalReversedVal + dictData(key)(month)("Reversed")
Next month
```

## Why This Works

1. **Separate Tracking**: By tracking Posted and Reversed amounts in separate dictionary entries, they no longer cancel each other out during data collection.

2. **Complete Visibility**: Both Posted and Reversed values are now visible in the output, even when they net to zero.

3. **Accurate Totals**: The Total column correctly sums all Posted amounts and all Reversed amounts separately.

4. **Correct Balance**: The Balance row still correctly shows Posted + Reversed, which may be zero when amounts cancel out.

## Files Modified

1. **provision.vba** (lines 114-126, 211-235, 322-326):
   - Changed dictionary structure to track Posted/Reversed separately
   - Updated output logic to read separate values
   - Updated Summary sheet logic

## Use Case Examples

### Example 1: Same Month Posted and Reversed
**Transactions:**
- August: +1000 (Posted)
- August: -1000 (Reversed)

**Output:**
```
Profit Center | Type     | Aug-24  | Total
10120002      | Posted   | 1000    | 1000
10120002      | Reversed | -1000   | -1000
10120002      | Balance  | 0       | 0
```

### Example 2: Multiple Postings and Reversals
**Transactions:**
- August: +1000 (Posted)
- August: +500 (Posted)
- August: -800 (Reversed)
- August: -200 (Reversed)

**Output:**
```
Profit Center | Type     | Aug-24  | Total
10120002      | Posted   | 1500    | 1500
10120002      | Reversed | -1000   | -1000
10120002      | Balance  | 500     | 500
```

### Example 3: Different Months
**Transactions:**
- August: +1000 (Posted)
- September: -1000 (Reversed)

**Output:**
```
Profit Center | Type     | Aug-24 | Sep-24 | Total
10120002      | Posted   | 1000   |        | 1000
10120002      | Reversed |        | -1000  | -1000
10120002      | Balance  | 1000   | -1000  | 0
```

## Verification Steps

To verify the fix works correctly:

1. Prepare test data with:
   - Same GL Account and Profit Center
   - Mix of positive (PostingKey 50) and negative (PostingKey 40) amounts
   - Some in the same month, some in different months

2. Run the `BuildProvisionReports` macro

3. Check each GL sheet to verify:
   - Posted row shows sum of all positive amounts for each month
   - Reversed row shows sum of all negative amounts for each month
   - Balance row shows Posted + Reversed (may be zero)
   - Total column correctly sums across all months

4. Check the Summary sheet to verify:
   - Posted column shows total of all positive amounts
   - Reversed column shows total of all negative amounts
   - Balance column shows Posted + Reversed

## Impact

- **Data Visibility**: Users can now see actual Posted and Reversed amounts instead of just the net result
- **Audit Trail**: Better visibility for understanding what provisions were posted and subsequently reversed
- **Data Accuracy**: No loss of information due to premature netting of amounts
- **User Experience**: Matches user expectation to see "actual posted and reversed values rather showing zero"
