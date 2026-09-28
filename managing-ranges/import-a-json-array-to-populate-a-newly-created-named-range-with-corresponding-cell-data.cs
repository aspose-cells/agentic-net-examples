// Title: How to import a JSON array into a named range in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that reads a JSON array, extracts headers and rows, fills an Aspose.Cells worksheet, creates a named range covering the data, and saves the workbook. | Show a step‑by‑step example of deserializing JSON to a list of dictionaries and populating an Aspose.Cells range with a custom name in C#.
// Common Searches: C# Aspose.Cells example to convert JSON array to Excel named range | how to create a named range from JSON data using Aspose.Cells .NET | populate Excel cells from a list of dictionaries with Aspose.Cells C# | Aspose.Cells deserialize JSON and set range name programmatically | write JSON records to Excel and define a named range in C#
// Tags: Aspose.Cells populate named range from JSON | C# create named range in Excel using Aspose.Cells | deserialize JSON to worksheet cells Aspose.Cells | write dictionary data to Excel range C# | Aspose.Cells range creation with dynamic data

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample creates a new workbook, deserializes a JSON array into a list of dictionaries, writes the dictionary keys as header cells and each record's values into subsequent rows, defines a named range that covers the populated area, and saves the file as JsonData.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample JSON array
            string json = @"
            [
                { ""Name"": ""John"", ""Age"": 30, ""City"": ""New York"" },
                { ""Name"": ""Anna"", ""Age"": 25, ""City"": ""London"" },
                { ""Name"": ""Mike"", ""Age"": 40, ""City"": ""Sydney"" }
            ]";

            // Deserialize JSON to a list of dictionaries
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            List<Dictionary<string, JsonElement>> records = JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(json, options);

            if (records == null || records.Count == 0)
                return;

            // Determine column headers from the first record
            var headers = new List<string>(records[0].Keys);
            int startRow = 0;      // zero‑based index (A1)
            int startColumn = 0;   // zero‑based index

            // Write headers
            for (int col = 0; col < headers.Count; col++)
            {
                sheet.Cells[startRow, startColumn + col].PutValue(headers[col]);
            }

            // Write data rows
            for (int row = 0; row < records.Count; row++)
            {
                var record = records[row];
                for (int col = 0; col < headers.Count; col++)
                {
                    string key = headers[col];
                    if (record.TryGetValue(key, out JsonElement value))
                    {
                        // Handle different JSON value types
                        switch (value.ValueKind)
                        {
                            case JsonValueKind.Number:
                                if (value.TryGetInt32(out int intVal))
                                    sheet.Cells[startRow + 1 + row, startColumn + col].PutValue(intVal);
                                else if (value.TryGetDouble(out double dblVal))
                                    sheet.Cells[startRow + 1 + row, startColumn + col].PutValue(dblVal);
                                break;
                            case JsonValueKind.String:
                                sheet.Cells[startRow + 1 + row, startColumn + col].PutValue(value.GetString());
                                break;
                            case JsonValueKind.True:
                            case JsonValueKind.False:
                                sheet.Cells[startRow + 1 + row, startColumn + col].PutValue(value.GetBoolean());
                                break;
                            default:
                                sheet.Cells[startRow + 1 + row, startColumn + col].PutValue(value.ToString());
                                break;
                        }
                    }
                }
            }

            // Define the range that includes headers and data
            int totalRows = records.Count + 1; // +1 for header row
            int totalCols = headers.Count;
            AsposeRange namedRange = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalCols);
            namedRange.Name = "MyJsonData";

            // Save the workbook (ensure the directory exists)
            string outputPath = "JsonData.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
