// Title: How to export a named range from an Excel workbook to a JSON array of cell addresses and values using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export named range to JSON file with cell addresses | How to retrieve cell address and value from a named range using Aspose.Cells in .NET | Save Excel named range as a JSON array with Aspose.Cells C# example | Aspose.Cells .NET get all cells in a named range and write to JSON | Export Excel named range data to JSON using Aspose.Cells library
// Tags: Aspose.Cells named range JSON export | C# Aspose.Cells range cell address extraction | Aspose.Cells serialize range values to JSON | Excel named range to JSON with .NET | Aspose.Cells iterate named range cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using System.Text.Json;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, accesses a specified named range, iterates through each cell to collect its address and value, serializes the collection into a formatted JSON array, and writes the output to a file while handling missing files and I/O errors.
class ExportNamedRangeToJson
{
    static void Main()
    {
        // Paths
        string excelPath = "input.xlsx";
        string jsonPath = "output.json";
        string namedRangeName = "MyRange";

        try
        {
            // Verify input file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file not found: {excelPath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(excelPath);

            // Retrieve named range
            Name namedRange = workbook.Worksheets.Names[namedRangeName];
            if (namedRange == null)
            {
                Console.WriteLine($"Named range '{namedRangeName}' not found.");
                return;
            }

            // Get the range object
            AsposeRange range = namedRange.GetRange();

            // Worksheet that contains the range
            Worksheet ws = range.Worksheet;

            // Collect cell address/value pairs
            var cellsInfo = new List<Dictionary<string, object>>();

            for (int row = range.FirstRow; row < range.FirstRow + range.RowCount; row++)
            {
                for (int col = range.FirstColumn; col < range.FirstColumn + range.ColumnCount; col++)
                {
                    Cell cell = ws.Cells[row, col];

                    var cellData = new Dictionary<string, object>
                    {
                        { "Address", cell.Name },
                        { "Value", cell.Value }
                    };

                    cellsInfo.Add(cellData);
                }
            }

            // Serialize to formatted JSON
            string json = JsonSerializer.Serialize(cellsInfo, new JsonSerializerOptions { WriteIndented = true });

            // Write JSON output safely
            try
            {
                File.WriteAllText(jsonPath, json);
                Console.WriteLine($"Named range '{namedRangeName}' exported to '{jsonPath}'.");
            }
            catch (Exception writeEx)
            {
                Console.WriteLine($"Failed to write JSON file: {writeEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
