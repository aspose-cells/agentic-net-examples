// Title: Validate that JSON exported from an Excel worksheet matches the worksheet row count using Aspose.Cells for .NET
// Common Searches: how to verify row count after exporting an Excel worksheet to JSON with Aspose.Cells in C# | compare Excel worksheet MaxDataRow with JSON array length using Aspose.Cells .NET | C# code to validate that exported JSON rows equal original Excel data rows
// Tags: Aspose.Cells export worksheet to JSON | row count validation between Excel and JSON | System.Text.Json parse JSON array length in C# | JsonSaveOptions usage with Aspose.Cells | compare Excel data rows with JSON rows

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Json;

namespace AsposeCellsJsonRowValidator
{
    // The example loads the first worksheet of an Excel file, determines the number of data rows using MaxDataRow, exports the used range to JSON with Aspose.Cells, parses the JSON with System.Text.Json to count rows, and validates that the JSON row count matches the original worksheet row count.
    class Program
    {
        static int Main(string[] args)
        {
            // Input Excel file – change the path as needed or pass as first argument
            string excelPath = args.Length > 0 ? args[0] : "input.xlsx";
            // Output JSON file (optional, for inspection)
            string jsonPath = "exported.json";

            if (!File.Exists(excelPath))
            {
                Console.Error.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                return 1;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);
                // Use the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Determine the number of rows that actually contain data
                // MaxDataRow is zero‑based; add 1 to get a count
                int dataRowCount = worksheet.Cells.MaxDataRow + 1;
                Console.WriteLine($"Worksheet '{worksheet.Name}' contains {dataRowCount} data rows.");

                // Create a range that covers the used area of the worksheet
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxCol = worksheet.Cells.MaxDataColumn;
                Aspose.Cells.Range usedRange = worksheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);

                // Export the range to a JSON string
                string json;
                try
                {
                    JsonSaveOptions jsonOptions = new JsonSaveOptions();
                    json = usedRange.ToJson(jsonOptions);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error exporting range to JSON: {ex.Message}");
                    return 3;
                }

                // Optionally write JSON to a file for manual inspection
                try
                {
                    File.WriteAllText(jsonPath, json);
                    Console.WriteLine($"Exported JSON written to '{jsonPath}'.");
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error writing JSON to file: {ex.Message}");
                }

                // Parse the JSON to count rows
                int jsonRowCount = CountRowsInExportedJson(json);
                Console.WriteLine($"JSON export contains {jsonRowCount} rows.");

                // Validate
                if (jsonRowCount == dataRowCount)
                {
                    Console.WriteLine("Validation succeeded: row counts match.");
                    return 0;
                }
                else
                {
                    Console.Error.WriteLine($"Validation failed: worksheet rows = {dataRowCount}, JSON rows = {jsonRowCount}.");
                    return 2;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
                return 3;
            }
        }

        private static int CountRowsInExportedJson(string json)
        {
            try
            {
                using JsonDocument doc = JsonDocument.Parse(json);
                JsonElement root = doc.RootElement;

                // If the root is an array, each element represents a row
                if (root.ValueKind == JsonValueKind.Array)
                {
                    return root.GetArrayLength();
                }

                // If the root is an object, look for a "Rows" array property
                if (root.ValueKind == JsonValueKind.Object &&
                    root.TryGetProperty("Rows", out JsonElement rowsElement) &&
                    rowsElement.ValueKind == JsonValueKind.Array)
                {
                    return rowsElement.GetArrayLength();
                }

                Console.Error.WriteLine("Warning: JSON does not contain a recognizable rows array. Assuming zero rows.");
                return 0;
            }
            catch (JsonException je)
            {
                Console.Error.WriteLine($"Error parsing JSON: {je.Message}");
                return 0;
            }
        }
    }
}
