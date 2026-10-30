// Title: Export the underlying data of an Excel PivotTable to a JSON array using Aspose.Cells for .NET (C#)
// Common Searches: how to extract pivot table source range and save as JSON with Aspose.Cells C# | Aspose.Cells C# export pivot table underlying data to a JSON file | read Excel pivot table data source and convert to JSON array using .NET | C# code to get pivot table source data address with Aspose.Cells | serialize Excel range to indented JSON using System.Text.Json and Aspose.Cells
// Tags: Aspose.Cells export pivot table source to JSON | C# read pivot table data source range | serialize Excel range as JSON array | extract underlying data from PivotTable .NET | Aspose.Cells get pivot table data source address | write JSON file from Excel data C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;          // For PivotTable
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range
using System.Text.Json;

// The example loads an Excel workbook, locates the first PivotTable, retrieves its data source address, reads the source range into a two‑dimensional array, builds a list of dictionaries using the header row as keys, serializes the collection to an indented JSON array with System.Text.Json, and writes the result to a file while handling missing files and I/O errors.
class PivotTableExport
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "pivot_underlying_data.json";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the PivotTable is on the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure a PivotTable exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found in the worksheet.");
                return;
            }

            // Get the first PivotTable
            PivotTable pivot = worksheet.PivotTables[0];

            // Retrieve the source data range address (e.g., "Sheet1!A1:D100")
            // In some Aspose.Cells versions DataSource returns a string array.
            string dataSourceAddress;
            if (pivot.DataSource is string[] sourceArray && sourceArray.Length > 0)
            {
                dataSourceAddress = sourceArray[0];
            }
            else
            {
                dataSourceAddress = pivot.DataSource?.ToString();
            }

            if (string.IsNullOrEmpty(dataSourceAddress))
            {
                Console.WriteLine("PivotTable data source is empty.");
                return;
            }

            // Split into sheet name and range address
            string[] parts = dataSourceAddress.Split('!');
            if (parts.Length != 2)
            {
                Console.WriteLine("Invalid data source address format.");
                return;
            }

            string sourceSheetName = parts[0].Trim('\'');
            string sourceRangeAddress = parts[1];

            // Get the worksheet containing the source data
            Worksheet sourceSheet = workbook.Worksheets[sourceSheetName];
            if (sourceSheet == null)
            {
                Console.WriteLine($"Source worksheet not found: {sourceSheetName}");
                return;
            }

            // Create the source range
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange(sourceRangeAddress);

            // Read values into a 2‑dimensional array
            object[,] rawValues = sourceRange.Value as object[,];
            if (rawValues == null)
            {
                Console.WriteLine("Unable to read source data.");
                return;
            }

            int rows = rawValues.GetLength(0);
            int cols = rawValues.GetLength(1);

            // Extract column headers from the first row
            List<string> headers = new List<string>();
            for (int c = 0; c < cols; c++)
            {
                headers.Add(rawValues[0, c]?.ToString() ?? $"Column{c + 1}");
            }

            // Build a list of dictionaries for each data row
            List<Dictionary<string, object>> jsonRows = new List<Dictionary<string, object>>();
            for (int r = 1; r < rows; r++) // skip header row
            {
                var rowDict = new Dictionary<string, object>();
                for (int c = 0; c < cols; c++)
                {
                    rowDict[headers[c]] = rawValues[r, c];
                }
                jsonRows.Add(rowDict);
            }

            // Serialize to JSON with indentation
            var options = new JsonSerializerOptions { WriteIndented = true };
            string jsonOutput = JsonSerializer.Serialize(jsonRows, options);

            // Write JSON to file (ensure directory exists)
            try
            {
                File.WriteAllText(outputPath, jsonOutput);
                Console.WriteLine($"PivotTable underlying data exported to JSON successfully: {outputPath}");
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
