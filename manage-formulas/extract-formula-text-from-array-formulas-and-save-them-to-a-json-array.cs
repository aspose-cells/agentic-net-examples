// Title: Extract Excel array formulas to a JSON file with Aspose.Cells for .NET
// Description: Loads an .xlsx workbook using Aspose.Cells, walks through every worksheet and used cell, detects cells flagged as array formulas, gathers their formula strings, serializes the collection into a pretty‑printed JSON array, and writes the output to a specified path.
// Keywords: Aspose.Cells C# array formula extraction | export Excel formulas to JSON | Cell.IsArrayFormula example | iterate worksheets Aspose.Cells | serialize formula list .NET
// Common Searches: how to retrieve array formula text with Aspose.Cells | save extracted Excel formulas as JSON in C# | list all array formulas in a workbook using Aspose
// Developer Intent: Collect every array‑formula expression from an Excel file and persist it as a JSON array.
// Use Cases: Audit all array calculations across a workbook for compliance reporting. | Create a portable reference of formulas for migration to another calculation engine. | Supply downstream analytics scripts with a JSON‑based catalog of spreadsheet logic.
// AI Prompts: Generate C# code that reads an Excel workbook, extracts array formulas, and writes them to a JSON file using Aspose.Cells. | Explain how to include each cell's address alongside its formula in the JSON output. | Show how to filter extracted formulas by worksheet name before serialization.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Loads an .xlsx workbook using Aspose.Cells, walks through every worksheet and used cell, detects cells flagged as array formulas, gathers their formula strings, serializes the collection into a pretty‑printed JSON array, and writes the output to a specified path.
class ExtractArrayFormulas
{
    static void Main()
    {
        // Path to the source Excel workbook
        string inputPath = "input.xlsx";

        // Path where the JSON array will be saved
        string outputPath = "arrayFormulas.json";

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(inputPath);

        // List to collect formula strings from array-formula cells
        List<string> arrayFormulas = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range limits
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Scan every cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Check if the cell contains an array formula
                    if (cell.IsArrayFormula)
                    {
                        // Store the formula text
                        arrayFormulas.Add(cell.Formula);
                    }
                }
            }
        }

        // Convert the list of formulas to a JSON array string
        string json = JsonSerializer.Serialize(arrayFormulas, new JsonSerializerOptions { WriteIndented = true });

        // Write the JSON string to the output file (save rule)
        File.WriteAllText(outputPath, json);

        Console.WriteLine($"Extracted {arrayFormulas.Count} array formulas and saved to '{outputPath}'.");
    }
}
