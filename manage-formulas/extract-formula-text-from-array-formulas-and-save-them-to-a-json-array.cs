// Title: C# – Extract Excel array formulas to a JSON file using Aspose.Cells
// Description: Load an Excel workbook with Aspose.Cells, scan each worksheet for cells where IsArrayFormula is true, collect the formula strings, and serialize them to a formatted JSON array saved to disk.
// Keywords: Aspose.Cells C# extract array formulas | export Excel formulas to JSON .NET | cell.IsArrayFormula example | C# write JSON file from Excel | Aspose.Cells workbook iteration | GitHub Aspose.Cells sample | Excel array formula extraction | JSON serialization System.Text.Json
// Common Searches: how to get array formulas from Excel using Aspose.Cells | save extracted formulas as JSON in C# | Aspose.Cells IsArrayFormula property usage | C# convert Excel formulas to JSON array | example code for extracting Excel formulas with Aspose
// Developer Intent: Retrieve every array‑formula string from an Excel workbook and output them as a JSON array.
// Use Cases: Generate a catalog of all array formulas in a financial model for documentation. | Migrate spreadsheet logic by exporting array formulas to JSON for analysis or conversion to another platform. | Audit workbooks for unexpected array formulas by comparing the exported list with a compliance whitelist.
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, extracts only array formulas, and returns them as a List<string>. | Modify the sample to include each cell address (e.g., A1) alongside its array formula in the JSON output. | Explain performance‑optimizing techniques for extracting array formulas from large workbooks with Aspose.Cells, including memory‑management tips.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Load an Excel workbook with Aspose.Cells, scan each worksheet for cells where IsArrayFormula is true, collect the formula strings, and serialize them to a formatted JSON array saved to disk.
class ExtractArrayFormulas
{
    static void Main()
    {
        // Path to the source Excel workbook
        string excelPath = "input.xlsx";

        // Path where the JSON array will be saved
        string jsonPath = "arrayFormulas.json";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(excelPath);

        // List to collect formula strings from array-formula cells
        List<string> arrayFormulas = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range to limit iteration
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Scan every cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell contains an array formula
                    if (cell.IsArrayFormula)
                    {
                        // Extract the formula text and add it to the list
                        arrayFormulas.Add(cell.Formula);
                    }
                }
            }
        }

        // Convert the list of formulas to a JSON array string
        string json = JsonSerializer.Serialize(arrayFormulas, new JsonSerializerOptions { WriteIndented = true });

        // Save the JSON string to a file (lifecycle rule: save)
        File.WriteAllText(jsonPath, json);

        Console.WriteLine($"Extracted {arrayFormulas.Count} array formulas and saved to '{jsonPath}'.");
    }
}
