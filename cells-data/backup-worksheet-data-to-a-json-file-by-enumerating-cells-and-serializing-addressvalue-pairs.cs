// Title: Export a worksheet’s cell address‑value pairs to a JSON file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export used range cells to JSON file | How to backup Excel worksheet data as JSON with Aspose.Cells | Serialize Excel cell addresses and values to JSON in .NET | Aspose.Cells iterate through worksheet cells and write JSON | Save non‑empty cells from Excel to JSON using C# Aspose.Cells
// Tags: Aspose.Cells export worksheet to JSON | C# enumerate used range cells Aspose.Cells | serialize Excel cell address value dictionary | backup Excel data as JSON .NET | write JSON file from Aspose.Cells workbook

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace WorksheetBackup
{
    // The example loads a workbook, selects the first worksheet, obtains its used range, iterates over each non‑empty cell, stores the cell address (e.g., "A1") and value in a dictionary, serializes the dictionary to an indented JSON string with System.Text.Json, ensures the output directory exists, and writes the JSON to a specified file while handling missing files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path
            string excelPath = @"C:\Data\SourceWorkbook.xlsx";

            // Output JSON file path
            string jsonPath = @"C:\Data\WorksheetBackup.json";

            try
            {
                // Verify that the source workbook exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Input file not found: {excelPath}");
                    return;
                }

                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(excelPath);

                // Choose the worksheet to backup (first worksheet in this example)
                Worksheet sheet = workbook.Worksheets[0];

                // Get the used range of the worksheet; may be null for an empty sheet
                var maxRange = sheet.Cells.MaxDisplayRange;
                if (maxRange == null)
                {
                    Console.WriteLine("The selected worksheet contains no data.");
                    return;
                }

                // Dictionary to hold cell address and its value
                var cellData = new Dictionary<string, object>();

                // Enumerate through each cell in the used range
                for (int row = maxRange.FirstRow; row <= maxRange.FirstRow + maxRange.RowCount - 1; row++)
                {
                    for (int col = maxRange.FirstColumn; col <= maxRange.FirstColumn + maxRange.ColumnCount - 1; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        // Skip empty cells if desired
                        if (cell.Value == null) continue;

                        // Store address (e.g., "A1") and its value
                        cellData[cell.Name] = cell.Value;
                    }
                }

                // Serialize the dictionary to JSON
                var jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                string jsonString = JsonSerializer.Serialize(cellData, jsonOptions);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(jsonPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write JSON to file
                File.WriteAllText(jsonPath, jsonString);

                Console.WriteLine($"Worksheet data has been backed up to {jsonPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
