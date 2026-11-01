// Title: Import a JSON array into an Excel worksheet as a table with automatic column mapping using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells import JSON array and create table with dynamic columns | Aspose.Cells .NET read JSON and populate worksheet rows automatically | Generate Excel column headers from JSON properties using Aspose.Cells | Save JSON data to an Excel file with Aspose.Cells in C# | Parse JsonDocument and write to Excel table using Aspose.Cells
// Tags: Aspose.Cells JSON to Excel conversion | automatic column mapping from JSON | create Excel table from JsonDocument | C# write JSON array to worksheet | dynamic Excel headers based on JSON fields

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The program creates a new workbook, parses a JSON array, derives column headers from the first object's properties, writes the headers to the first row, iterates through each JSON element to populate subsequent rows handling numbers, strings, and booleans, and saves the workbook as 'JsonData.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Load / Create Workbook ----------
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // ---------- Sample JSON ----------
            // This JSON string represents an array of objects.
            // In a real scenario, you might read it from a file or other source.
            string json = @"
            [
                { ""Id"": 1, ""Name"": ""Alice"", ""Score"": 85.5 },
                { ""Id"": 2, ""Name"": ""Bob"",   ""Score"": 92.0 },
                { ""Id"": 3, ""Name"": ""Carol"", ""Score"": 78.0 }
            ]";

            // ---------- Parse JSON ----------
            using JsonDocument doc = JsonDocument.Parse(json);
            JsonElement root = doc.RootElement;

            if (root.ValueKind != JsonValueKind.Array)
            {
                Console.WriteLine("JSON root is not an array.");
                return;
            }

            // Determine column headers from the first element's property names
            List<string> headers = new List<string>();
            if (root.GetArrayLength() > 0)
            {
                foreach (JsonProperty prop in root[0].EnumerateObject())
                {
                    headers.Add(prop.Name);
                }
            }

            // Write headers to the first row (A1, B1, ...)
            for (int i = 0; i < headers.Count; i++)
            {
                sheet.Cells[0, i].PutValue(headers[i]);
            }

            // Write data rows starting from row index 1
            int currentRow = 1;
            foreach (JsonElement element in root.EnumerateArray())
            {
                for (int col = 0; col < headers.Count; col++)
                {
                    string header = headers[col];
                    if (element.TryGetProperty(header, out JsonElement value))
                    {
                        // Handle different JSON value types
                        switch (value.ValueKind)
                        {
                            case JsonValueKind.Number:
                                if (value.TryGetInt32(out int intVal))
                                    sheet.Cells[currentRow, col].PutValue(intVal);
                                else if (value.TryGetDouble(out double dblVal))
                                    sheet.Cells[currentRow, col].PutValue(dblVal);
                                break;
                            case JsonValueKind.String:
                                sheet.Cells[currentRow, col].PutValue(value.GetString());
                                break;
                            case JsonValueKind.True:
                            case JsonValueKind.False:
                                sheet.Cells[currentRow, col].PutValue(value.GetBoolean());
                                break;
                            default:
                                sheet.Cells[currentRow, col].PutValue(value.ToString());
                                break;
                        }
                    }
                    else
                    {
                        // If the property is missing, leave the cell blank
                        sheet.Cells[currentRow, col].PutValue(string.Empty);
                    }
                }
                currentRow++;
            }

            // ---------- Save Workbook ----------
            string outputPath = "JsonData.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to convert zero‑based column index to Excel column name (A, B, ..., AA, AB, ...)
    static string CellIndexToName(int columnIndex)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string result = string.Empty;
        int dividend = columnIndex + 1;

        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            result = letters[modulo] + result;
            dividend = (dividend - modulo) / 26;
        }

        return result;
    }
}
