// Title: Create an Excel table from a JSON array with automatic column generation using Aspose.Cells for .NET
// Common Searches: how to read a JSON file and populate an Excel table with Aspose.Cells in C# | generate Excel table columns dynamically from JSON properties using Aspose.Cells | C# Aspose.Cells import JSON array into worksheet as ListObject | save JSON data to a formatted XLSX table with Aspose.Cells .NET | apply a predefined table style to a JSON‑driven Excel table in C#
// Tags: automatic column creation from JSON Aspose.Cells | create ListObject from JSON array C# | apply table style to Excel table Aspose.Cells | import JSON data to XLSX using Aspose.Cells | auto‑populate worksheet headers from JSON properties

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example reads a JSON file containing an array, extracts the property names of the first element to build column headers, writes those headers and all row values into a new workbook worksheet, adds a ListObject (Excel table) over the populated range, applies a built‑in table style, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the JSON source file
            string jsonPath = "data.json";

            // Verify that the JSON file exists to avoid FileNotFoundException
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"JSON file not found: {jsonPath}");
                return;
            }

            // Load JSON content
            string jsonContent = File.ReadAllText(jsonPath);

            // Parse the JSON document
            using JsonDocument doc = JsonDocument.Parse(jsonContent);
            JsonElement root = doc.RootElement;

            if (root.ValueKind != JsonValueKind.Array)
            {
                Console.WriteLine("The JSON root element must be an array.");
                return;
            }

            // Determine column names from the first object in the array
            List<string> columns = new List<string>();
            foreach (JsonProperty prop in root[0].EnumerateObject())
            {
                columns.Add(prop.Name);
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Write header row
            for (int c = 0; c < columns.Count; c++)
            {
                sheet.Cells[0, c].PutValue(columns[c]);
            }

            // Write data rows
            int rowIndex = 1; // start after header
            foreach (JsonElement element in root.EnumerateArray())
            {
                for (int c = 0; c < columns.Count; c++)
                {
                    string colName = columns[c];
                    if (element.TryGetProperty(colName, out JsonElement value))
                    {
                        switch (value.ValueKind)
                        {
                            case JsonValueKind.Number:
                                if (value.TryGetInt64(out long l))
                                    sheet.Cells[rowIndex, c].PutValue(l);
                                else if (value.TryGetDouble(out double d))
                                    sheet.Cells[rowIndex, c].PutValue(d);
                                break;
                            case JsonValueKind.String:
                                sheet.Cells[rowIndex, c].PutValue(value.GetString());
                                break;
                            case JsonValueKind.True:
                            case JsonValueKind.False:
                                sheet.Cells[rowIndex, c].PutValue(value.GetBoolean());
                                break;
                            case JsonValueKind.Null:
                                sheet.Cells[rowIndex, c].PutValue(string.Empty);
                                break;
                            default:
                                sheet.Cells[rowIndex, c].PutValue(value.ToString());
                                break;
                        }
                    }
                    else
                    {
                        sheet.Cells[rowIndex, c].PutValue(string.Empty);
                    }
                }
                rowIndex++;
            }

            // Define the range for the table (including header row)
            int totalRows = rowIndex - 1; // number of data rows
            int totalCols = columns.Count;
            string lastColLetter = CellsHelper.ColumnIndexToName(totalCols - 1);
            string tableRange = $"A1:{lastColLetter}{totalRows + 1}"; // +1 because rows are 1‑based in Excel notation

            // Add a ListObject (Excel table) to the worksheet
            int tableIndex = sheet.ListObjects.Add("JsonDataTable", tableRange, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "JsonDataTable";
            table.TableStyleType = TableStyleType.TableStyleMedium9;
            table.ShowTableStyleFirstColumn = true;
            table.ShowTableStyleLastColumn = true;

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
