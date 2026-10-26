// Title: Create a JSON backup of an Excel worksheet by iterating cells with Aspose.Cells for .NET (C#)
// Common Searches: C# code to export Excel worksheet cells to a JSON file using Aspose.Cells | How to iterate over the used range in Aspose.Cells and save cell address/value pairs as JSON | Backup Excel data to JSON with the Aspose.Cells .NET library | Serialize Excel cell addresses and values to JSON in a C# console application | Aspose.Cells example for converting a worksheet to a JSON dictionary
// Tags: Aspose.Cells iterate used range | Excel worksheet to JSON conversion | cell address value mapping C# | Aspose.Cells JSON serialization example | backup Excel data as JSON .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// This C# console sample loads an Excel workbook with Aspose.Cells, selects the first worksheet, determines its used range, iterates each populated cell, stores the cell's address (e.g., "A1") and value in a dictionary, serializes the dictionary to an indented JSON string using System.Text.Json, and writes the JSON to a file, handling missing files and runtime exceptions.
class WorksheetBackup
{
    static void Main()
    {
        // Paths for input Excel file and output JSON backup
        string excelPath = "input.xlsx";
        string jsonPath = "backup.json";

        try
        {
            // Verify that the source Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            // Load the workbook from the file system
            Workbook workbook = new Workbook(excelPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Dictionary to hold cell address‑value pairs
            var cellData = new Dictionary<string, object>();

            // Access the cells collection of the worksheet
            Cells cells = sheet.Cells;

            // Determine the used range (last row/column with data)
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Enumerate each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Skip cells without a value to keep the JSON concise
                    if (cell.Value == null)
                        continue;

                    // Use the cell's address (e.g., "A1") as the key
                    string address = cell.Name;

                    // Store the cell's value; Aspose.Cells returns the value as an object
                    object value = cell.Value;

                    cellData[address] = value;
                }
            }

            // Serialize the dictionary to a formatted JSON string
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(cellData, jsonOptions);

            // Write the JSON content to the specified file
            File.WriteAllText(jsonPath, json);

            Console.WriteLine($"Backup completed successfully. JSON saved to \"{jsonPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
