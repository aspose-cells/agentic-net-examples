// Title: Identify and export Excel cells that use IFERROR or IFNA formulas with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that scans a workbook and returns a list of cell addresses whose formulas contain IFERROR or IFNA. | Extend the example to also capture formulas with ISERROR and write the collected results to a CSV file via Aspose.Cells. | Create a reusable method that accepts a Worksheet and returns all cells where the formula includes any error‑handling function (IFERROR, IFNA, ISERROR).
// Common Searches: C# Aspose.Cells find cells with IFERROR formula in an existing Excel file | how to list formulas that use IFNA using Aspose.Cells .NET | extract error handling functions from workbook formulas programmatically with Aspose.Cells | enumerate cells containing IFERROR or IFNA across all worksheets in a .xlsx using C#
// Tags: detect error‑handling functions in Excel formulas with Aspose.Cells | retrieve cells that use error functions via C# Aspose.Cells | enumerate formula cells across worksheets using Aspose.Cells | C# scan .xlsx for error functions in formulas | export identified error‑handling formulas to console

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The program loads an Excel workbook, iterates through each worksheet's used range, collects the addresses and formulas of cells that contain IFERROR or IFNA, and writes the list to the console.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        var workbook = new Workbook("input.xlsx");

        // List to collect cells that contain IFERROR or IFNA in their formulas
        var errorFormulaCells = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            var cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Scan the used range of the worksheet
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    var cell = cells[row, col];
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        string formula = cell.Formula;

                        // Detect IFERROR or IFNA (case‑insensitive)
                        if (formula.IndexOf("IFERROR", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            formula.IndexOf("IFNA", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Store the sheet name, cell address, and the formula
                            errorFormulaCells.Add($"{sheet.Name}!{cell.Name}: {formula}");
                        }
                    }
                }
            }
        }

        // Export the collected formulas (here we simply write them to the console)
        Console.WriteLine("Formulas containing IFERROR or IFNA:");
        foreach (var entry in errorFormulaCells)
        {
            Console.WriteLine(entry);
        }
    }
}
