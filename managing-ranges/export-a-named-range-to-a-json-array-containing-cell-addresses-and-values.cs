// Title: How to export a named range from an Excel file to a JSON array of cell addresses and values using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export defined range to JSON file | Get cell address and value from a named range using Aspose.Cells | Serialize Excel named range to JSON array in .NET | Aspose.Cells iterate over named range cells and output JSON | How to write named range data to JSON with Aspose.Cells C#
// Tags: Aspose.Cells export named range to JSON | Aspose.Cells get range by name C# | serialize Excel range to JSON .NET | named range cell iteration Aspose.Cells | cell address extraction Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, retrieves a defined named range, loops through each cell to capture its full address (including worksheet name) and value, stores these pairs in a list, serializes the list to a formatted JSON array, prints the JSON, and saves it to a file.
class ExportNamedRangeToJson
{
    static void Main()
    {
        // Path to the Excel file
        string excelPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
            return;
        }

        // Name of the defined range to export
        string rangeName = "MyNamedRange";

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Get the range by its defined name (use fully qualified type to avoid ambiguity)
            Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName(rangeName);
            if (namedRange == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            // Prepare a list to hold address/value pairs
            var cellsData = new List<Dictionary<string, object>>();

            // Iterate through each cell in the range
            for (int row = 0; row < namedRange.RowCount; row++)
            {
                for (int col = 0; col < namedRange.ColumnCount; col++)
                {
                    // Get the cell object
                    Cell cell = namedRange[row, col];

                    // Build a dictionary for the current cell
                    var cellInfo = new Dictionary<string, object>
                    {
                        // Full address including sheet name (e.g., Sheet1!A1)
                        { "Address", $"{cell.Worksheet.Name}!{cell.Name}" },

                        // The cell's value; use .Value to get the underlying object
                        { "Value", cell.Value }
                    };

                    cellsData.Add(cellInfo);
                }
            }

            // Serialize the list to JSON
            string json = JsonSerializer.Serialize(cellsData, new JsonSerializerOptions { WriteIndented = true });

            // Output JSON to console
            Console.WriteLine(json);

            // Save JSON to a file
            string jsonPath = "named_range.json";
            File.WriteAllText(jsonPath, json);
            Console.WriteLine($"JSON exported to {jsonPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
