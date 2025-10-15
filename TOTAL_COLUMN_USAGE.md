# Total Column Feature - User Guide

## Overview
The provision reports now include a "Total" column at the end of each sheet that shows cumulative amounts for each Profit Center across all months.

## What's New

### GL Detail Sheets
Each GL account sheet (e.g., "Provisions") now has:
- A "Total" column as the last column after all month columns
- Posted row shows: Sum of all positive amounts
- Reversed row shows: Sum of all negative amounts
- Balance row shows: Net position (Posted + Reversed)

### Summary Sheet
The Summary sheet has the same Total column showing cumulative amounts for each GL Account / Profit Center combination.

## Example

### Before (without Total column):
```
Profit Center | Type     | Aug-24 | Dec-24 | Jan-25 | Feb-25
10120008      | Posted   |        | 70000  |        |
10120008      | Reversed |        | -5000  |        | -75000
10120008      | Balance  |        | 65000  |        | -75000
```

### After (with Total column):
```
Profit Center | Type     | Aug-24 | Dec-24 | Jan-25 | Feb-25 | Total
10120008      | Posted   |        | 70000  |        |        | 70000
10120008      | Reversed |        | -5000  |        | -75000 | -80000
10120008      | Balance  |        | 65000  |        | -75000 | -10000
```

## Understanding the Total Column

### Posted Total
- Shows the sum of ALL positive amounts (provisions posted) across ALL months
- Represents the total amount of provisions created for this Profit Center

### Reversed Total
- Shows the sum of ALL negative amounts (provision reversals) across ALL months
- Represents the total amount of provisions that have been reversed
- Note: This will typically be a negative number

### Balance Total
- Shows the net position: Posted Total + Reversed Total
- Represents the current outstanding provision balance for this Profit Center
- Positive balance = Net provision liability
- Negative balance = Over-reversed (more reversals than postings)

## How to Use

### Quick Reconciliation
1. Look at the Posted Total to see total provisions created
2. Look at the Reversed Total to see total provisions reversed
3. Look at the Balance Total to see the net outstanding provision

### Identifying Issues
- If Balance Total is significantly different from expected, review individual months
- Large negative balances may indicate over-reversals that need investigation
- Compare Posted and Reversed totals to understand provision activity

### Month-by-Month Analysis
- Individual month columns still show the detailed activity
- Use month columns to understand when provisions were posted/reversed
- Use Total column to understand the cumulative impact

## Notes

- The Total column is automatically calculated - no manual entry needed
- Partial reversals are properly accumulated (added to existing values)
- Empty cells are treated as zero in calculations
- The Total column updates every time the macro is run

## Support

If you notice any discrepancies in the Total column or have questions about the calculations, please refer to:
- `TOTAL_COLUMN_FIX.md` for technical details
- Your accounting team for business rules
- The VBA code maintainer for calculation logic
