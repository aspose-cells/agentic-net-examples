// Title: Validate that the row count exported to JSON matches the original Excel worksheet using Aspose.Cells in C#
// AI Prompts: Write a C# console application that opens an Excel workbook with Aspose.Cells, determines the data row count of the first worksheet, saves that worksheet as a JSON file, parses the JSON to count the entries in the "Rows" array, and prints whether the two counts are equal. | Create a .NET method that takes an Excel file path, uses Aspose.Cells to export the first sheet to JSON, reads the generated JSON, compares the JSON row count to the worksheet's MaxDataRow+1, and returns a boolean indicating a match.
// Common Searches: how to compare Excel worksheet row count with exported JSON rows using Aspose.Cells C# | Aspose.Cells verify row count after saving sheet as JSON | C# count rows in Aspose.Cells JSON output | validate that JSON export from Excel retains same number of rows Aspose.Cells | check MaxDataRow against JSON Rows array Aspose.Cells .NET
// Tags: Aspose.Cells export worksheet to JSON | compare worksheet row count with JSON rows | validate JSON row count Aspose.Cells | C# MaxDataRow row count verification | JSON Rows array length Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Text.Json;
using System.Text.Json.Nodes;

// The example loads an Excel file, calculates the number of data rows in the first worksheet using MaxDataRow+1, saves that worksheet as a JSON file with Aspose.Cells, reads and parses the JSON to count the objects in the "Rows" array, and then compares the two counts, outputting a success or mismatch message while handling missing files and exceptions.
class JsonRowCountValidator
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Verify that the Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the number of rows that contain data in the worksheet
            // MaxDataRow is zero‑based, so add 1 to get the count
            int originalRowCount = sheet.Cells.MaxDataRow + 1;

            // Export the worksheet to JSON format
            string jsonPath = "output.json";

            // Save only the first worksheet as JSON
            workbook.Save(jsonPath, SaveFormat.Json);

            // Verify that the JSON file was created
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"Error: Failed to create JSON file at '{jsonPath}'.");
                return;
            }

            // Read the generated JSON file
            string jsonContent = File.ReadAllText(jsonPath);

            // Parse the JSON; Aspose.Cells writes an array of row objects under the "Rows" property
            JsonNode rootNode = JsonNode.Parse(jsonContent);
            JsonArray rowsArray = rootNode?["Rows"]?.AsArray();

            // Count the rows present in the exported JSON
            int jsonRowCount = rowsArray?.Count ?? 0;

            // Validate that the row counts match
            if (originalRowCount == jsonRowCount)
            {
                Console.WriteLine($"Success: Row count matches ({originalRowCount} rows).");
            }
            else
            {
                Console.WriteLine($"Mismatch: Original worksheet has {originalRowCount} rows, but JSON contains {jsonRowCount} rows.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
