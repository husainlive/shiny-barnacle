# Test Scenario: Summary Sheet Linked to GL Expense Sheets

## Test Objective
Verify that Summary sheet cells contain Excel formulas linking to the Total column cells in individual GL expense sheets.

## Test Data Setup

### Input Data (Sample CSV/Excel rows)
```
Document Date | Profit Center | Profit Center: Short Text | Posting Key | Company Code Currency Value | Offsetting Account
01/08/2024   | 10120008     | Cost Center A             | 50          | 1000                        | 400100
01/12/2024   | 10120008     | Cost Center A             | 50          | 70000                       | 400100
01/12/2024   | 10120008     | Cost Center A             | 40          | 5000                        | 400100
01/08/2024   | 10120003     | Cost Center B             | 50          | 1500                        | 400100
01/01/2025   | 10120003     | Cost Center B             | 50          | 2000                        | 400100
01/12/2024   | 10120003     | Cost Center B             | 40          | 200                         | 400100
```

### GL Mapping (in Personal.xlsb)
```
GL Code | Description
400100  | Provisions
```

## Expected Output

### GL Sheet: "Provisions"
```
A           B        C      D       E      F
Profit Ctr  Type     08-24  12-24   01-25  Total
10120008    Posted   1000   70000   0      71000
10120008    Reversed 0      -5000   0      -5000
10120008    Balance  1000   65000   0      66000

10120003    Posted   1500   0       2000   3500
10120003    Reversed 0      -200    0      -200
10120003    Balance  1500   -200    2000   3300
```

Note: The Total column (F) should contain actual values, not formulas.

### Summary Sheet
```
A           B                    C                       D                      
Profit Ctr  Provisions - Posted  Provisions - Reversed  Provisions - Balance
10120003    ='Provisions'!F5     ='Provisions'!F6       ='Provisions'!F7
10120008    ='Provisions'!F2     ='Provisions'!F3       ='Provisions'!F4
```

**Important:** The cells should display the formulas when viewed in the Excel formula bar.

When displayed as values, it should look like:
```
A           B                    C                       D                      
Profit Ctr  Provisions - Posted  Provisions - Reversed  Provisions - Balance
10120003    3500                 -200                   3300
10120008    71000                -5000                  66000
```

## Test Steps

### 1. Preparation
- [ ] Open Excel with the test data
- [ ] Ensure GL_Mapping sheet exists in Personal.xlsb with the mapping above
- [ ] Open VBA Editor (Alt+F11)
- [ ] Import provision.vba module

### 2. Execute Macro
- [ ] Run `BuildProvisionReports` macro
- [ ] Verify no compilation errors
- [ ] Verify no runtime errors

### 3. Verify GL Sheet Creation
- [ ] Check that "Provisions" sheet is created
- [ ] Verify header row: "Profit Center" | "Type" | month columns | "Total"
- [ ] Verify each Profit Center has 3 rows (Posted, Reversed, Balance)
- [ ] Verify all Profit Center values are filled in column A
- [ ] Verify Total column contains aggregated values (not formulas)

### 4. Verify Summary Sheet Creation
- [ ] Check that "Summary" sheet is created
- [ ] Verify header row: "Profit Center" | "Provisions - Posted" | "Provisions - Reversed" | "Provisions - Balance"
- [ ] Verify one row per Profit Center
- [ ] Verify Profit Centers are sorted alphabetically

### 5. Verify Summary Sheet Formulas (KEY TEST)
For each non-zero value in the Summary sheet:

#### Test Case 1: Provisions - Posted for 10120003
- [ ] Click on cell B2 (assuming 10120003 is in row 2)
- [ ] Check Excel formula bar
- [ ] **Expected**: `='Provisions'!F5` (or similar, depending on actual row)
- [ ] Navigate to Provisions sheet, cell F5
- [ ] **Expected**: Value should be 3500

#### Test Case 2: Provisions - Reversed for 10120003
- [ ] Click on cell C2
- [ ] Check Excel formula bar
- [ ] **Expected**: `='Provisions'!F6`
- [ ] Navigate to Provisions sheet, cell F6
- [ ] **Expected**: Value should be -200

#### Test Case 3: Provisions - Balance for 10120003
- [ ] Click on cell D2
- [ ] Check Excel formula bar
- [ ] **Expected**: `='Provisions'!F7`
- [ ] Navigate to Provisions sheet, cell F7
- [ ] **Expected**: Value should be 3300

#### Test Case 4: Provisions - Posted for 10120008
- [ ] Click on cell B3 (assuming 10120008 is in row 3)
- [ ] Check Excel formula bar
- [ ] **Expected**: `='Provisions'!F2`
- [ ] Navigate to Provisions sheet, cell F2
- [ ] **Expected**: Value should be 71000

### 6. Verify Live Update Feature
- [ ] In Provisions sheet, locate cell F2 (Posted Total for 10120008)
- [ ] Note current value (should be 71000)
- [ ] Change the value to 80000
- [ ] Switch to Summary sheet
- [ ] Check cell B3 (Provisions - Posted for 10120008)
- [ ] **Expected**: Value should automatically update to 80000
- [ ] Undo the change (Ctrl+Z) to restore original value

### 7. Verify Navigation Feature
- [ ] In Summary sheet, select cell with formula (e.g., B2)
- [ ] Press Ctrl+[ (navigate to precedent)
- [ ] **Expected**: Excel should jump to Provisions sheet, cell F5
- [ ] OR: With cell selected, press F5, type the reference shown in formula bar, press Enter
- [ ] **Expected**: Should navigate to the correct cell in Provisions sheet

### 8. Verify Column Letter Conversion
This tests the ColumnNumberToLetter function accuracy:

| Column Number | Expected Letter | Formula Example |
|---------------|----------------|-----------------|
| 1             | A              | ='Sheet'!A1     |
| 26            | Z              | ='Sheet'!Z1     |
| 27            | AA             | ='Sheet'!AA1    |
| 52            | AZ             | ='Sheet'!AZ1    |
| 53            | BA             | ='Sheet'!BA1    |

Based on the test data, the Total column should be:
- If there are 3 months + header + Total = column F (6th column)
- Check that formulas reference column F, not column 6

### 9. Edge Cases

#### Empty/Zero Values
- [ ] Verify that if Posted value is 0, no formula is created (cell is empty)
- [ ] Verify that if Reversed value is 0, no formula is created (cell is empty)
- [ ] Verify Balance formula is created only if balance is non-zero

#### Multiple GL Accounts
If test data includes multiple GL accounts (e.g., 400100 and 400200):
- [ ] Verify Summary sheet has columns for both GL accounts
- [ ] Verify formulas reference the correct GL sheet for each account
- [ ] Example: Provisions columns reference ='Provisions'!..., Equipment columns reference ='Equipment'!...

#### Special Characters in Sheet Names
If GL account names have spaces or special characters:
- [ ] Verify sheet names are enclosed in single quotes
- [ ] Example: `='Provisions and Accruals'!F2` (note the quotes)

## Success Criteria

All of the following must be true:
- ✅ Summary sheet cells contain formulas, not plain values
- ✅ Formulas reference the correct GL sheet name
- ✅ Formulas reference the correct cell (Total column, appropriate row for Posted/Reversed/Balance)
- ✅ Formula syntax is valid: `='SheetName'!ColumnRow`
- ✅ Values displayed in Summary match source values in GL sheets
- ✅ Changing a GL sheet value automatically updates Summary sheet
- ✅ No compilation or runtime errors
- ✅ All existing functionality preserved (3-row structure, sorting, etc.)

## Common Issues and Troubleshooting

### Issue: Formulas show as text (e.g., "='Provisions'!F2" as text)
- **Cause**: Cell format is set to Text
- **Fix**: Select cells, format as General, press F2 then Enter to re-evaluate

### Issue: #REF! error in Summary cells
- **Cause**: Referenced GL sheet or cell doesn't exist
- **Debug**: Check if GL sheet name matches formula, verify row/column numbers

### Issue: Values don't update when GL sheet changes
- **Cause**: Automatic calculation is disabled
- **Fix**: Go to Formulas tab > Calculation Options > Automatic

### Issue: Wrong column letter in formula
- **Cause**: ColumnNumberToLetter function error
- **Debug**: Check the Total column number stored in dictCellRefs, verify conversion

## Notes for Developers

- The dictCellRefs dictionary stores cell positions during GL sheet creation
- Key format: "GLDescription|ProfitCenter" (e.g., "Provisions|10120008")
- Values stored: SheetName, PostedRow, ReversedRow, BalanceRow, TotalCol
- ColumnNumberToLetter handles Excel's 1-based indexing correctly
- Formulas are created using .Formula property, not .Value property
