// Title: Verify that Excel formulas use English (US) function names and commas as separators with Aspose.Cells for .NET
// Common Searches: how to check Excel formula language and argument separators using Aspose.Cells C# | detect semicolon separators in workbook formulas with Aspose.Cells .NET | validate English function names in Excel formulas programmatically in C# | Aspose.Cells iterate cells and verify formula locale | C# code to ensure Excel formulas use commas not semicolons
// Tags: Aspose.Cells formula locale validation | C# verify Excel formula separators | English function name detection in formulas | comma argument separator check in .xlsx | regex formula pattern matching with Aspose.Cells

using Aspose.Cells;
using System;
using System.Text.RegularExpressions;

// The example loads an Excel workbook using Aspose.Cells, walks through each worksheet and cell, and for every cell that contains a formula it confirms that commas are used as argument separators, semicolons are absent, and the function names match the English (US) pattern via a regular expression. Cells with non‑conforming formulas are reported to the console, and the workbook can be saved after the verification step.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate through all cells that contain formulas
            foreach (Cell cell in cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    string formula = cell.Formula;

                    // Check for commas as argument separators
                    bool hasCommas = formula.Contains(",");

                    // Ensure there are no semicolons (used in some locales)
                    bool hasSemicolons = formula.Contains(";");

                    // Simple validation that function names are English (uppercase letters followed by '(')
                    bool englishFunctions = Regex.IsMatch(formula, @"[A-Z]{2,}\(");

                    // Report any formula that does not meet the criteria
                    if (!hasCommas || hasSemicolons || !englishFunctions)
                    {
                        Console.WriteLine($"Invalid formula in cell {cell.Name} (Sheet: {sheet.Name}): {formula}");
                    }
                }
            }
        }

        // Optionally save the workbook after verification
        workbook.Save("output.xlsx");
    }
}
