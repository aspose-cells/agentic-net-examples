// Title: Replace deprecated REPT function with REPEAT in all Excel formulas using Aspose.Cells for .NET
// Common Searches: how to replace REPT with REPEAT in Excel formulas using Aspose.Cells C# | bulk replace deprecated Excel functions in a workbook with Aspose.Cells | update formulas containing REPT across all worksheets Aspose.Cells .NET | programmatically change REPT to REPEAT in an Excel file using Aspose.Cells | C# code to migrate REPT function to REPEAT in an Aspose.Cells workbook
// Tags: replace REPT with REPEAT Aspose.Cells | bulk formula update .NET | Excel deprecated function migration C# | programmatic formula editing Aspose.Cells | update cell formulas across worksheets

// Load an existing workbook (replace with your actual file path)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Iterate through all worksheets
foreach (Aspose.Cells.Worksheet sheet in workbook.Worksheets)
{
    // Get the cells collection of the current worksheet
    Aspose.Cells.Cells cells = sheet.Cells;

    // Iterate through all used cells in the worksheet
    foreach (Aspose.Cells.Cell cell in cells)
    {
        // Process only cells that contain a formula
        if (cell.IsFormula)
        {
            string formula = cell.Formula;

            // Check if the formula uses the deprecated REPT function
            if (formula.IndexOf("REPT(", System.StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Replace REPT with REPEAT (preserve original case for the rest of the formula)
                string updatedFormula = System.Text.RegularExpressions.Regex.Replace(
                    formula,
                    @"\bREPT\b",
                    "REPEAT",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // Assign the updated formula back to the cell
                cell.Formula = updatedFormula;
            }
        }
    }
}

// Save the modified workbook (replace with your desired output path)
workbook.Save("output.xlsx");
