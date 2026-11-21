// Title: Create a separate Excel worksheet for each table in a JSON array using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells generate a new worksheet for each element of a JSON array of tables | How to import multiple JSON tables into separate sheets in an Excel workbook with Aspose.Cells | Read a JSON file containing an array of tables and export each table to its own sheet using Aspose.Cells in .NET | Aspose.Cells C# example for converting a JSON array of objects to a multi‑sheet Excel file | Map JSON object properties to column headers and rows in separate worksheets with Aspose.Cells
// Tags: Aspose.Cells JSON to multi‑sheet Excel conversion | C# create worksheets from JSON array | Aspose.Cells write JSON object properties to cells | Export JSON tables to separate Excel sheets | Dynamic worksheet generation from JSON in .NET

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsJsonToExcel
{
    // The sample reads a JSON file that contains a top‑level array where each element represents a table (an array of objects). For every table it adds a new worksheet to an Aspose.Cells Workbook, writes column headers derived from the first object's property names, fills subsequent rows with primitive values, and finally saves the workbook as an .xlsx file. The code includes validation of the input file, JSON structure checks, and automatic creation of the output directory.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "data.json";
                const string outputPath = "output.xlsx";

                // Verify that the input JSON file exists
                if (!File.Exists(inputPath))
                {
                    Console.Error.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Read the JSON content
                string jsonContent = File.ReadAllText(inputPath);

                // Parse the JSON; expect a top‑level array where each element is a table
                using JsonDocument doc = JsonDocument.Parse(jsonContent);
                JsonElement root = doc.RootElement;

                if (root.ValueKind != JsonValueKind.Array)
                {
                    Console.Error.WriteLine("The JSON file must contain an array of tables.");
                    return;
                }

                // Create a new workbook and clear the default worksheet
                Workbook workbook = new Workbook();
                workbook.Worksheets.Clear();

                int sheetNumber = 1;

                // Iterate over each table element in the array
                foreach (JsonElement tableElement in root.EnumerateArray())
                {
                    try
                    {
                        // Add a new worksheet for the current table
                        Worksheet sheet = workbook.Worksheets.Add($"Sheet{sheetNumber}");
                        Cells cells = sheet.Cells;

                        // Expect each table element to be an array of objects (rows)
                        if (tableElement.ValueKind != JsonValueKind.Array)
                        {
                            Console.Error.WriteLine($"Table {sheetNumber} is not an array. Skipping.");
                            continue;
                        }

                        JsonElement.ArrayEnumerator rows = tableElement.EnumerateArray();

                        // Determine column names from the first row
                        if (!rows.MoveNext())
                        {
                            Console.Error.WriteLine($"Table {sheetNumber} contains no rows. Skipping.");
                            continue;
                        }

                        JsonElement firstRow = rows.Current;
                        if (firstRow.ValueKind != JsonValueKind.Object)
                        {
                            Console.Error.WriteLine($"First row of table {sheetNumber} is not an object. Skipping.");
                            continue;
                        }

                        // Extract column names
                        var columnNames = new System.Collections.Generic.List<string>();
                        foreach (JsonProperty prop in firstRow.EnumerateObject())
                        {
                            columnNames.Add(prop.Name);
                        }

                        // Write header
                        for (int col = 0; col < columnNames.Count; col++)
                        {
                            cells[0, col].PutValue(columnNames[col]);
                        }

                        // Write first row data
                        WriteRow(cells, 1, columnNames, firstRow);

                        int currentRowIndex = 2; // start after header and first data row

                        // Write remaining rows
                        while (rows.MoveNext())
                        {
                            WriteRow(cells, currentRowIndex, columnNames, rows.Current);
                            currentRowIndex++;
                        }

                        sheetNumber++;
                    }
                    catch (Exception exTable)
                    {
                        Console.Error.WriteLine($"Error processing table {sheetNumber}: {exTable.Message}");
                    }
                }

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to an Excel file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to write a JSON object into a worksheet row
        private static void WriteRow(Cells cells, int rowIndex, System.Collections.Generic.List<string> columnNames, JsonElement rowElement)
        {
            for (int col = 0; col < columnNames.Count; col++)
            {
                string colName = columnNames[col];
                if (rowElement.TryGetProperty(colName, out JsonElement valueElement))
                {
                    // Simplified handling for primitive JSON types
                    switch (valueElement.ValueKind)
                    {
                        case JsonValueKind.String:
                            cells[rowIndex, col].PutValue(valueElement.GetString());
                            break;
                        case JsonValueKind.Number:
                            if (valueElement.TryGetInt64(out long l))
                                cells[rowIndex, col].PutValue(l);
                            else if (valueElement.TryGetDouble(out double d))
                                cells[rowIndex, col].PutValue(d);
                            break;
                        case JsonValueKind.True:
                        case JsonValueKind.False:
                            cells[rowIndex, col].PutValue(valueElement.GetBoolean());
                            break;
                        case JsonValueKind.Null:
                            cells[rowIndex, col].PutValue(string.Empty);
                            break;
                        default:
                            // For complex types, store the raw JSON text
                            cells[rowIndex, col].PutValue(valueElement.GetRawText());
                            break;
                    }
                }
                else
                {
                    // Column missing in this row; leave cell empty
                    cells[rowIndex, col].PutValue(string.Empty);
                }
            }
        }
    }
}
