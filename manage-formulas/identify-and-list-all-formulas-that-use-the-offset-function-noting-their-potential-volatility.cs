// Title: How to enumerate OFFSET formulas and determine their volatility in an Excel workbook with Aspose.Cells for .NET
// Common Searches: C# code to find all OFFSET functions in an Excel file using Aspose.Cells | list volatile formulas containing OFFSET in a .xlsx workbook with Aspose.Cells .NET | detect and output cells that use the OFFSET function in a spreadsheet via Aspose.Cells | retrieve formula text and volatility flag for OFFSET formulas in C# | iterate through worksheets and identify volatile OFFSET formulas using Aspose.Cells API
// Tags: enumerate OFFSET formulas Aspose.Cells | detect volatile Excel functions C# | list workbook formulas containing OFFSET | extract formula text Aspose.Cells .NET | identify volatile functions in .xlsx with Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an Excel workbook, walks through each worksheet and every cell, checks whether a cell contains a formula, and if the formula text includes the OFFSET function (case‑insensitive). For each match it prints the sheet name, cell address, the full formula (prefixed with '=') and marks the formula as volatile.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Loop through all cells that contain formulas
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula)
                {
                    string formula = cell.Formula; // Formula text without the leading '='

                    // Check if the formula uses the OFFSET function (case‑insensitive)
                    if (formula.IndexOf("OFFSET", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // OFFSET is a volatile function, so we note its volatility
                        Console.WriteLine(
                            $"Sheet: {sheet.Name}, Cell: {cell.Name}, Formula: ={formula}, Volatile: Yes");
                    }
                }
            }
        }
    }
}
