// Title: Read a JSON array and automatically generate an Excel table with dynamic columns using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a JSON file containing an array of objects, determines column names from the first object's properties, populates an Aspose.Cells worksheet with headers and rows, creates a ListObject covering the data range, and saves the workbook as an .xlsx file. | Enhance the sample to apply a built‑in table style, auto‑fit column widths, and set a custom display name for the table when converting JSON data to an Excel table with Aspose.Cells.
// Common Searches: C# Aspose.Cells create Excel table from JSON array with dynamic columns | How to map JSON object properties to Excel headers using Aspose.Cells | Aspose.Cells ListObject generation from JSON data in .NET | Automatically size columns after importing JSON into Excel with Aspose.Cells | Save JSON data as .xlsx file with table formatting using Aspose.Cells
// Tags: json to excel table Aspose.Cells | auto generate worksheet headers from json Aspose.Cells | excel table object creation from json Aspose.Cells | apply built‑in table style Aspose.Cells | auto fit column widths Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The program reads a JSON file containing an array of objects, extracts the property names of the first element to build column headers, writes the headers and each object's values into a new worksheet, creates a ListObject covering the populated range, optionally styles the table, and saves the workbook as JsonData.xlsx using Aspose.Cells.
class Program
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

            // Load JSON content
            string jsonContent = File.ReadAllText(jsonPath);
            JsonDocument doc = JsonDocument.Parse(jsonContent);
            JsonElement root = doc.RootElement;

            if (root.ValueKind != JsonValueKind.Array)
            {
                Console.WriteLine("JSON root must be an array of objects.");
                return;
            }

            // Determine column names from the first object
            List<string> columns = new List<string>();
            foreach (JsonProperty prop in root[0].EnumerateObject())
            {
                columns.Add(prop.Name);
            }

            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // Write header row
            for (int c = 0; c < columns.Count; c++)
            {
                ws.Cells[0, c].PutValue(columns[c]);
            }

            // Write data rows
            int rowIndex = 1;
            foreach (JsonElement element in root.EnumerateArray())
            {
                for (int col = 0; col < columns.Count; col++)
                {
                    string colName = columns[col];
                    if (element.TryGetProperty(colName, out JsonElement value))
                    {
                        switch (value.ValueKind)
                        {
                            case JsonValueKind.Number:
                                if (value.TryGetInt64(out long l))
                                    ws.Cells[rowIndex, col].PutValue(l);
                                else if (value.TryGetDouble(out double d))
                                    ws.Cells[rowIndex, col].PutValue(d);
                                break;
                            case JsonValueKind.String:
                                ws.Cells[rowIndex, col].PutValue(value.GetString());
                                break;
                            case JsonValueKind.True:
                            case JsonValueKind.False:
                                ws.Cells[rowIndex, col].PutValue(value.GetBoolean());
                                break;
                            case JsonValueKind.Null:
                                ws.Cells[rowIndex, col].PutValue(string.Empty);
                                break;
                            default:
                                ws.Cells[rowIndex, col].PutValue(value.ToString());
                                break;
                        }
                    }
                    else
                    {
                        ws.Cells[rowIndex, col].PutValue(string.Empty);
                    }
                }
                rowIndex++;
            }

            // Define the range that includes header and data
            int totalRows = rowIndex; // rowIndex points to the row after the last data row
            int totalCols = columns.Count;
            string tableRange = $"A1:{CellIndexToName(totalRows - 1, totalCols - 1)}";

            // Add a ListObject (Excel table) to the worksheet
            int tableIdx = ws.ListObjects.Add("JsonDataTable", tableRange, true);
            ListObject table = ws.ListObjects[tableIdx];
            table.DisplayName = "JsonDataTable";
            table.ShowTableStyleFirstColumn = true;
            table.ShowTableStyleLastColumn = true;
            // Optional: set a built‑in table style if desired
            // table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Save the workbook
            string outputPath = "JsonData.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to convert zero‑based row/column indexes to Excel cell address (e.g., 0,0 -> A1)
    static string CellIndexToName(int row, int col)
    {
        string colName = "";
        int dividend = col + 1;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            colName = Convert.ToChar('A' + modulo) + colName;
            dividend = (dividend - modulo) / 26;
        }
        return $"{colName}{row + 1}";
    }
}
