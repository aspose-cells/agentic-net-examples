// Title: Export Excel formulas to JSON by worksheet using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook with Aspose.Cells, scans each worksheet’s used range, captures cells where IsFormula is true, converts each cell to JSON (address, formula, style), groups the results by sheet name, and writes an indented JSON file.
// Keywords: Aspose.Cells | C# | extract Excel formulas | JSON export | worksheet formulas | Cell.ToJson | used range iteration | .NET | workbook automation | formula extraction
// Common Searches: Aspose.Cells extract formulas C# | export Excel formulas as JSON .NET | list formula cells with addresses using Aspose | convert workbook formulas to JSON file | C# code to get all formulas from Excel | Cell.ToJson example Aspose
// Developer Intent: Retrieve every formula cell from an Excel file and write a JSON document that maps each worksheet name to an array of cell objects.
// Use Cases: Document spreadsheet logic by exporting formulas with their locations for review. | Create JSON snapshots of workbook formulas to compare different versions. | Feed extracted formulas into a validation service that checks for prohibited functions or references.
// AI Prompts: Generate C# code using Aspose.Cells that extracts all formula cells from a workbook and outputs a JSON file mapping worksheet names to cell objects. | Show how to modify the example to include each cell’s calculated value alongside its formula in the JSON output. | Explain alternative serialization methods for the formula dictionary without relying on System.Text.Json.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsFormulaExtractor
{
    // Loads an Excel workbook with Aspose.Cells, scans each worksheet’s used range, captures cells where IsFormula is true, converts each cell to JSON (address, formula, style), groups the results by sheet name, and writes an indented JSON file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Path where the resulting JSON will be saved
            string jsonOutputPath = "formulas.json";

            // Load the workbook (creation/loading rule)
            Workbook workbook = new Workbook(excelPath);

            // Dictionary to hold sheet name -> list of cell JSON strings (each contains address and formula)
            var sheetFormulas = new Dictionary<string, List<string>>();

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // List to collect JSON representations of formula cells in the current sheet
                var formulaCells = new List<string>();

                Cells cells = sheet.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // Scan the used range for formula cells
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell != null && cell.IsFormula)
                        {
                            // Use the provided Cell.ToJson method to obtain JSON for the cell
                            string cellJson = cell.ToJson();

                            // Add the cell JSON to the list
                            formulaCells.Add(cellJson);
                        }
                    }
                }

                // If the sheet contains any formula cells, add them to the result dictionary
                if (formulaCells.Count > 0)
                {
                    sheetFormulas[sheet.Name] = formulaCells;
                }
            }

            // Serialize the complete structure to a JSON string
            // The structure is: { "SheetName": [ "{cell json}", "{cell json}", ... ], ... }
            string finalJson = JsonSerializer.Serialize(sheetFormulas, new JsonSerializerOptions { WriteIndented = true });

            // Save the JSON document (saving rule)
            File.WriteAllText(jsonOutputPath, finalJson);

            Console.WriteLine($"Formulas extracted and saved to '{jsonOutputPath}'.");
        }
    }
}
