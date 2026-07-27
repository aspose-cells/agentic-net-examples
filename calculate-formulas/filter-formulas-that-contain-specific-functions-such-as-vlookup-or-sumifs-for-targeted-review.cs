// Title: C# Example: Filter Cells Containing VLOOKUP or SUMIFS Formulas with Aspose.Cells
// Description: Loads an Excel workbook, scans the first worksheet for formulas that include VLOOKUP or SUMIFS using Aspose.Cells FindOptions (OnlyFormulas + Contains), gathers each unique cell, prints its address and formula, and saves the workbook unchanged. Ideal for .NET developers needing to audit or extract specific functions from Excel files.
// Keywords: Aspose.Cells | C# | .NET | find formulas | VLOOKUP | SUMIFS | filter cells by function | Excel workbook analysis | FindOptions OnlyFormulas | code example | GitHub
// Common Searches: Aspose.Cells find VLOOKUP formulas C# | search for SUMIFS cells using Aspose.Cells .NET | list Excel cells that contain specific functions | avoid duplicate matches when searching multiple formulas | C# example to audit lookup functions in a workbook
// Developer Intent: Locate and list every cell whose formula includes VLOOKUP or SUMIFS.
// Use Cases: Create an audit report of all lookup and conditional‑sum formulas in a spreadsheet. | Validate that prohibited functions are not present before distributing an Excel file. | Extract formula strings for bulk refactoring or migration to newer functions.
// AI Prompts: Generate C# code with Aspose.Cells that finds all INDEX function formulas and outputs their addresses. | Show how to replace each VLOOKUP formula with an XLOOKUP equivalent using Aspose.Cells. | Provide a snippet that logs matched cells to a CSV file while keeping the original workbook unchanged.

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

class FilterFormulas
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Functions we want to locate in formulas
        string[] targetFunctions = { "VLOOKUP", "SUMIFS" };

        // Store cells that contain any of the target functions
        List<Cell> matchedCells = new List<Cell>();

        // Scan all used cells in the worksheet
        foreach (Cell cell in cells)
        {
            if (cell.IsFormula)
            {
                string formula = cell.Formula;
                foreach (string func in targetFunctions)
                {
                    if (formula.IndexOf(func, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        matchedCells.Add(cell);
                        break; // No need to check other functions for this cell
                    }
                }
            }
        }

        // Output the addresses and formulas of matched cells
        Console.WriteLine("Cells containing VLOOKUP or SUMIFS:");
        foreach (Cell c in matchedCells)
        {
            Console.WriteLine($"{c.Name}: {c.Formula}");
        }

        // Optional: highlight matched cells for visual review
        foreach (Cell c in matchedCells)
        {
            Style style = c.GetStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            c.SetStyle(style);
        }

        // Save the workbook with highlights (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}
