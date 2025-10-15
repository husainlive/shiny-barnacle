# Test Scenario for Fix #4

## Problem Statement
The individual expenses output was showing only "Posted" rows for each Profit Center, with all amounts accumulated into that single row, instead of showing the proper 3-row structure (Posted, Reversed, Balance).

## Root Cause
When an existing Profit Center was found in the GL sheet (lines 156-161), the code would skip setting the Type labels for the Reversed and Balance rows. This caused those rows to appear empty or missing in the output.

## Expected Behavior After Fix

### Input Data Example:
- Profit Center: 10120002
- GL Code: 123456
- Transactions:
  - Aug-24: PostingKey=50, Amount=1000 (becomes +1000)
  - Aug-24: PostingKey=40, Amount=1000 (becomes -1000)
  - Oct-24: PostingKey=50, Amount=1000 (becomes +1000)
  - Nov-24: PostingKey=50, Amount=1000 (becomes +1000)
  - Nov-24: PostingKey=40, Amount=1000 (becomes -1000)
  - Dec-24: PostingKey=50, Amount=1000 (becomes +1000)

### Expected Output in GL Sheet:
```
Profit Center | Type     | Aug-24 | Oct-24 | Nov-24 | Dec-24
10120002      | Posted   | 1000   | 1000   | 1000   | 1000
              | Reversed | -1000  | 0      | -1000  | 0
              | Balance  | 0      | 1000   | 0      | 1000
```

### How the Fix Works:
1. When processing Profit Center 10120002 for the first time:
   - Creates row 2 with Profit Center value in column A
   - Sets "Posted" in B2, "Reversed" in B3, "Balance" in B4

2. When processing Profit Center 10120002 for subsequent months:
   - Finds existing row 2 with matching Profit Center
   - **With the fix**: Still ensures "Posted", "Reversed", and "Balance" labels are set in B2, B3, B4
   - Continues to fill in the month data in the appropriate rows

3. Data population:
   - Positive amounts (from PostingKey=50) go to row 2 (Posted)
   - Negative amounts (from PostingKey=40) go to row 3 (Reversed)
   - Balance in row 4 is calculated as row 2 + row 3

## Testing Steps:
1. Create a test Excel file with sample data containing:
   - Multiple transactions for the same Profit Center
   - Mix of PostingKey 50 (positive) and 40 (negative)
   - Multiple months
2. Run the BuildProvisionReports macro
3. Verify each GL sheet shows:
   - Three rows per Profit Center (Posted, Reversed, Balance)
   - Posted row contains only positive amounts
   - Reversed row contains only negative amounts
   - Balance row equals Posted + Reversed for each month
4. Verify the Summary sheet also shows the 3-row structure for each GL/PC combination
