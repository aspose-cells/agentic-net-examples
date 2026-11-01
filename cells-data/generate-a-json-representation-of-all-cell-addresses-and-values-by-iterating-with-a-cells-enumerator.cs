// Title: Export non‑empty Excel cell addresses and values to indented JSON using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, iterates over every cell in the first worksheet via the Cells enumerator, and builds a JSON array where each element holds the cell's address (e.g., "A1") and its stringified value. | Extend the program to also capture each cell's data type (Number, String, DateTime, etc.) and include it in the JSON objects. | Refactor the solution so that the JSON string is returned from a method instead of being written to a file, making it suitable for an ASP.NET Core API response. | Add a command‑line option to exclude cells whose values are only whitespace when creating the JSON output.
// Common Searches: Aspose.Cells C# how to export worksheet cells to a JSON file | Iterate over cells with Aspose.Cells and get cell address and value | Serialize non‑empty Excel cells to JSON in .NET Core | C# convert Excel data to JSON including cell coordinates using Aspose.Cells | Skip empty or blank cells when exporting Excel to JSON with Aspose.Cells
// Tags: Aspose.Cells export cells to JSON | enumerate worksheet cells Aspose.Cells | filter non‑empty cells Aspose.Cells | cell address serialization C# | Excel to JSON conversion Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    // Loads an Excel workbook with Aspose.Cells, enumerates all cells in the first worksheet, collects the address and string representation of each non‑null cell into a list, serializes the list to indented JSON, and writes the result to a file named cells.json.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a list to hold cell address/value pairs
            var cellsData = new List<object>();

            // Iterate over all cells using the Cells enumerator
            foreach (Cell cell in worksheet.Cells)
            {
                // Skip empty cells if desired
                if (cell.Value == null) continue;

                // Add an anonymous object with Address and Value
                cellsData.Add(new
                {
                    Address = cell.Name,          // e.g., "A1"
                    Value = cell.Value.ToString() // Convert value to string for JSON
                });
            }

            // Serialize the list to JSON
            string json = JsonSerializer.Serialize(cellsData, new JsonSerializerOptions
            {
                WriteIndented = true
            });

            // Save JSON to a file (replace with your desired output path)
            string outputPath = "cells.json";
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"Cell data exported to {outputPath}");
        }
    }
}
