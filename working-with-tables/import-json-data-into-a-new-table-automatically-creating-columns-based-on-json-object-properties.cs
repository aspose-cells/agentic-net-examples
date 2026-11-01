// Title: Create an Excel table from a JSON array with dynamic columns using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a JSON file containing an array of objects, automatically creates matching column headers, fills an Aspose.Cells worksheet, adds an Excel table via the ListObjects API, and saves the workbook. | Extend the example to detect nested JSON objects and export each nesting level to its own worksheet while preserving the table styling applied by Aspose.Cells. | Add a command‑line option that lets the user choose an Aspose.Cells table style and resize columns to fit their content after the data is written.
// Common Searches: asp.net read json array and export to excel table using aspose.cells | c# dynamic column creation from json properties for excel workbook | how to add an excel table after loading json data with Aspose.Cells in C# | resize columns to fit content after json import to excel c#
// Tags: json-to-excel table generation Aspose.Cells | dynamic column mapping from JSON properties | excel table creation Aspose.Cells | auto-fit columns Aspose.Cells | custom table style selection Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The program reads a JSON array from a file, extracts property names to generate header columns, writes each object's values into an Aspose.Cells worksheet, creates an Excel ListObject table with a predefined style, auto‑fits the columns, and saves the result as output.xlsx.
class JsonToExcelTable
{
    static void Main()
    {
        try
        {
            // Path to the JSON file
            string jsonPath = "data.json";

            // Verify that the JSON file exists
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"JSON file not found: {jsonPath}");
                return;
            }

            // Read JSON content
            string jsonContent = File.ReadAllText(jsonPath);

            // Parse JSON (expects an array of objects)
            using JsonDocument doc = JsonDocument.Parse(jsonContent);
            JsonElement root = doc.RootElement;

            if (root.ValueKind != JsonValueKind.Array)
            {
                Console.WriteLine("JSON root must be an array of objects.");
                return;
            }

            // Collect column names from the first object
            List<string> columnNames = new List<string>();
            foreach (JsonProperty prop in root[0].EnumerateObject())
            {
                columnNames.Add(prop.Name);
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Write header row
            for (int col = 0; col < columnNames.Count; col++)
            {
                sheet.Cells[0, col].PutValue(columnNames[col]);
            }

            // Write data rows
            int rowIndex = 1; // start after header
            foreach (JsonElement element in root.EnumerateArray())
            {
                for (int col = 0; col < columnNames.Count; col++)
                {
                    string columnName = columnNames[col];
                    if (element.TryGetProperty(columnName, out JsonElement valueElement))
                    {
                        // Handle different JSON value types
                        switch (valueElement.ValueKind)
                        {
                            case JsonValueKind.Number:
                                if (valueElement.TryGetInt64(out long longVal))
                                    sheet.Cells[rowIndex, col].PutValue(longVal);
                                else if (valueElement.TryGetDouble(out double doubleVal))
                                    sheet.Cells[rowIndex, col].PutValue(doubleVal);
                                break;
                            case JsonValueKind.String:
                                sheet.Cells[rowIndex, col].PutValue(valueElement.GetString());
                                break;
                            case JsonValueKind.True:
                            case JsonValueKind.False:
                                sheet.Cells[rowIndex, col].PutValue(valueElement.GetBoolean());
                                break;
                            case JsonValueKind.Null:
                                sheet.Cells[rowIndex, col].PutValue(string.Empty);
                                break;
                            default:
                                // For complex types, store the raw JSON string
                                sheet.Cells[rowIndex, col].PutValue(valueElement.GetRawText());
                                break;
                        }
                    }
                    else
                    {
                        // Property missing in this object; leave cell empty
                        sheet.Cells[rowIndex, col].PutValue(string.Empty);
                    }
                }
                rowIndex++;
            }

            // Define the range for the table (including header)
            int totalRows = rowIndex; // rowIndex is already one past the last data row
            int totalCols = columnNames.Count;

            // Add a ListObject (Excel Table) over the range
            int tableIndex = sheet.ListObjects.Add(0, 0, totalRows - 1, totalCols - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set table properties (use DisplayName instead of Name for compatibility)
            table.DisplayName = "JsonDataTable";
            table.ShowTableStyleFirstColumn = true;
            table.ShowTableStyleLastColumn = true;
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Auto-fit columns for better appearance
            sheet.AutoFitColumns();

            // Save the workbook
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
