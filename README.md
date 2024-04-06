**Excel Data Pivoting Macro**

**Description:**
This VBA macro enables the transformation of data in an Excel worksheet by splitting values in a single row into multiple rows, effectively pivoting the data. Each value in the original row, separated by a semicolon, is distributed across multiple rows in a target worksheet while maintaining the corresponding ID for each value.

**Functionality:**
1. **Set Worksheets:** Defines the source worksheet (`Sheet1`) containing the original data and creates or references a target worksheet (`Result`) where the pivoted data will be placed.
2. **Clear Existing Data:** Clears any existing data in the target worksheet to ensure a clean slate for the pivoted data.
3. **Add Headers:** Adds headers to the target worksheet to identify the columns of the pivoted data.
4. **Loop Through Rows:** Iterates through each row in the source worksheet, starting from the second row.
5. **Split Values:** Splits values in each cell of the current row by a semicolon, creating an array of individual values.
6. **Populate Target Worksheet:** For each value in the array, adds a new row to the target worksheet with the corresponding ID from the source row and the value placed in the appropriate column.
7. **Increment Counter:** Increments the row counter for the target worksheet to ensure the next set of values is placed in the correct row.
8. **Repeat for Each Column:** The above steps are repeated for each column in the source worksheet, allowing for the pivoting of multiple sets of values.

**Usage:**
1. Open the Excel workbook containing the data to be pivoted.
2. Press `Alt + F11` to access the Visual Basic for Applications (VBA) editor.
3. Insert a new module and paste the provided VBA macro code.
4. Customize the macro according to your data layout (e.g., specify the source and target worksheet names, adjust column ranges, and headers).
5. Save the workbook as a macro-enabled Excel file (.xlsm) to preserve the VBA code.
6. Run the macro from the Excel workbook by pressing `Alt + F8` and selecting the `SplitAndPivotValues` macro.

**Example:**
```vba
Sub SplitAndPivotValues()
    ' VBA macro code goes here...
End Sub
```

**Notes:**
- Ensure that the source and target worksheets are correctly specified in the code.
- Adjust column ranges, headers, and other parameters as needed based on your specific data layout.
- This macro is useful for transforming data where multiple values are combined in a single row into a structured format suitable for analysis and reporting.

If you encounter any issues or have questions, feel free to reach out for assistance!

**License:**
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

**Author:**
Rakkesh R

**Contact:**
rakkesh30.mbm@gmail.com

**Repository:**
