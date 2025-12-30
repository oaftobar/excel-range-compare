Create a Python program that compares two specified ranges of cells 
in a spreadsheet, and identifies the differences between them.

1. Load the Spreadsheet:
    Load the "Employee Sales" workbook into a python program.

2. Take User Input to Specify Worksheets and Ranges:
    Ask the user to specify two worksheets & corresponding cell ranges within WB
    User provides:
        a. Worksheet Names: first & second WS. EXACT NAME!
        b. Cell Ranges: "A1:C9" EXACT MATCH!
        c. (Optional) Check if Worksheets & Cell Ranges have same dimensions.
        Error message if not, loop until input matches

3. Perform Comparison Between Specified Cell Ranges:
    a. Iterating Over Rows & Columns: row by row, and within each, cell by cell.
    b. Comparing Cell Values: check whether the value in the cell from first range is the value in second range.
    Direct comparison of the two values; numbers, strings, or some other data type.
    c. Recording Differences: record differences "programatically" (i.e. list of lists)
        i. Which row it occurred in
        ii. Which column (using the column header)
        iii. The differing vallue from the first range
        iv. The differing value from the second range
    The program will have examined every pair of cells and noted any differences.

4. Create New Worksheet and Write Differences
    a. Programmatically create a new WS named "Diffs"
    b. Programmatically write the following column headers to cells A1:E1 of the "Diffs" WS:
        i. "Row"
        ii. "Column"
        iii. "Value 1"
        iv. "Value 2"
        v. "Delta" (Only needed if you're completing the optional Step 5 below)
    c. For each difference recorded (Step 3) write a row in this new WS.        

5. OPTIONAL: Calculate Percentage Differences for Numeric Values
    Write formulas to the "Delta" column of the "Diffs" WS which calculate ABSOLUTE DIFFERENCE between "Value 1" and "Value 2"
    ONLY performed if the values are both numbers
    =ABS(value1, value2)/AVERAGE(value2, value2)

6. OPTIONAL: Format the "Delta" Column:
    Text color should be red
    The values should also be formatted as percentages