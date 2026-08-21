// Title: Export All Excel Formulas to JSON with Aspose.Cells for .NET (C#)
// Description: A C# example that loads an Excel workbook using Aspose.Cells, scans every worksheet and used cell, captures the sheet name, cell address, and formula text, then serializes the collection into a pretty‑printed JSON file (formulas.json) for external analysis.
// Keywords: Aspose.Cells export formulas JSON | C# extract Excel formulas | list workbook formulas Aspose | serialize Excel formulas to JSON | Aspose.Cells iterate worksheets cells
// Common Searches: how to extract all formulas from an Excel file using Aspose.Cells C# | save Excel formulas as JSON with Aspose.Cells | export workbook formula list to JSON .NET | Aspose.Cells iterate cells to get formulas
// Developer Intent: Retrieve every formula in a workbook and write the details to a JSON file.
// Use Cases: Create an audit log of all formulas for compliance or review. | Feed the JSON output into a data‑science pipeline to analyze formula complexity. | Generate documentation that maps each sheet and cell to its underlying calculation.
// AI Prompts: Generate C# code that adds the calculated value and cell style to the JSON export. | Enhance the sample with error handling for missing files and atomic write operations. | Show how to filter and export only formulas that contain specific functions such as VLOOKUP or SUMIF.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsFormulaExport
{
    // A C# example that loads an Excel workbook using Aspose.Cells, scans every worksheet and used cell, captures the sheet name, cell address, and formula text, then serializes the collection into a pretty‑printed JSON file (formulas.json) for external analysis.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Prepare a collection to hold formula information
            var formulas = new List<Dictionary<string, string>>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all used cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Check if the cell contains a formula
                    if (cell.IsFormula)
                    {
                        // Store sheet name, cell name (e.g., "A1") and the formula text
                        var entry = new Dictionary<string, string>
                        {
                            { "Sheet", sheet.Name },
                            { "Cell", cell.Name },
                            { "Formula", cell.Formula }
                        };
                        formulas.Add(entry);
                    }
                }
            }

            // Serialize the list of formulas to a formatted JSON string
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string jsonOutput = JsonSerializer.Serialize(formulas, jsonOptions);

            // Save the JSON string to a file
            string outputPath = "formulas.json";
            File.WriteAllText(outputPath, jsonOutput);

            Console.WriteLine($"Exported {formulas.Count} formulas to '{outputPath}'.");
        }
    }
}
