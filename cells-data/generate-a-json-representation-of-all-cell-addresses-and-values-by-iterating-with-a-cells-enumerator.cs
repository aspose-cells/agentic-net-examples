// Title: Generate a JSON file of all non‑empty cell addresses and values from an Excel worksheet using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells iterate worksheet cells and export to JSON | how to get cell address and value with Aspose.Cells enumerator | save non empty Excel cells as JSON using System.Text.Json | Aspose.Cells load workbook and list populated cells in JSON format | export Excel data to JSON file with Aspose.Cells C# example
// Tags: Aspose.Cells enumerate non‑empty cells | JSON export of Excel cell data C# | cell address and value extraction Aspose.Cells | System.Text.Json serialization of worksheet data | load workbook and iterate cells Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example loads an Excel workbook, accesses the first worksheet, and uses the Cells enumerator to walk through every cell. For each cell that contains a value, it records the cell's address (e.g., "A1") and its value in a dictionary. All dictionaries are collected into a list, which is then serialized into a nicely indented JSON string with System.Text.Json and written to an output file. The code also includes basic file‑existence checking and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.json";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook (add password to LoadOptions if needed)
            var loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Collect non‑empty cell data
            var cellsData = new List<Dictionary<string, object>>();
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.Value == null) continue;

                var cellInfo = new Dictionary<string, object>
                {
                    { "Address", cell.Name }, // e.g., "A1"
                    { "Value",   cell.Value } // actual cell value
                };
                cellsData.Add(cellInfo);
            }

            // Serialize to formatted JSON
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(cellsData, jsonOptions);

            // Write JSON to output file
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"JSON representation of cells has been saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any errors, including password protection or file issues
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
