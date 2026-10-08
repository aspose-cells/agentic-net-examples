// Title: Create separate worksheets for each table in a JSON array using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that reads a JSON file containing an array of table objects, adds a new worksheet for each table with Aspose.Cells, and writes the table's data into the sheet. | Modify the code to set each worksheet's name from the table's "Name" property (or a generated name) and save the workbook as an XLSX file.
// Common Searches: how to import multiple JSON tables into separate Excel worksheets using Aspose.Cells C# | aspocells read json array and create a worksheet per element | c# map json Data array to cells with Aspose.Cells | dynamic worksheet naming from json property using Aspose.Cells .NET | save workbook as xlsx after processing json tables in C#
// Tags: json-array-to-multiple-worksheets aspocells | populate-worksheet-cells System.Text.Json | worksheet-naming-from-json aspocells | save-workbook-xlsx aspocells | csharp-iterate-json-add-worksheets

using System;
using System.IO;
using Aspose.Cells;
using System.Text.Json;

// The example reads a JSON file (tables.json) that contains an array of table objects, each with an optional "Name" and a two‑dimensional "Data" array. It creates a new Aspose.Cells workbook, clears the default sheet, and for each table adds a worksheet named from the "Name" field or a generated name. The code iterates through the rows and columns of the "Data" array, writes values of various JSON types into cells using PutValue, and finally saves the workbook as Output.xlsx. Error handling ensures missing files or malformed JSON are reported.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the JSON file that contains an array of tables
            string jsonPath = "tables.json";

            // Verify that the JSON file exists
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"Error: JSON file not found at path '{jsonPath}'.");
                return;
            }

            // Load the entire JSON content as a string
            string jsonContent = File.ReadAllText(jsonPath);

            // Parse the JSON content
            using JsonDocument doc = JsonDocument.Parse(jsonContent);
            JsonElement root = doc.RootElement;

            if (root.ValueKind != JsonValueKind.Array)
            {
                Console.WriteLine("Error: Expected a JSON array at the root.");
                return;
            }

            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Remove the default worksheet that Aspose.Cells creates
            workbook.Worksheets.Clear();

            // Iterate through each table object in the JSON array
            foreach (JsonElement table in root.EnumerateArray())
            {
                // Determine the worksheet name; if not provided, generate a default name
                string sheetName;
                if (table.TryGetProperty("Name", out JsonElement nameProp) && nameProp.ValueKind == JsonValueKind.String)
                {
                    sheetName = nameProp.GetString();
                }
                else
                {
                    sheetName = $"Sheet{workbook.Worksheets.Count + 1}";
                }

                // Add a new worksheet to the workbook and set its name
                Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet.Name = sheetName;

                // Retrieve the 2‑dimensional data array for the current table
                if (!table.TryGetProperty("Data", out JsonElement rows) || rows.ValueKind != JsonValueKind.Array)
                {
                    // Skip if there is no data array
                    continue;
                }

                int rowIndex = 0;
                foreach (JsonElement row in rows.EnumerateArray())
                {
                    if (row.ValueKind != JsonValueKind.Array)
                    {
                        rowIndex++;
                        continue;
                    }

                    int colIndex = 0;
                    foreach (JsonElement cell in row.EnumerateArray())
                    {
                        object value = null;
                        switch (cell.ValueKind)
                        {
                            case JsonValueKind.String:
                                value = cell.GetString();
                                break;
                            case JsonValueKind.Number:
                                if (cell.TryGetInt64(out long l))
                                    value = l;
                                else if (cell.TryGetDouble(out double d))
                                    value = d;
                                break;
                            case JsonValueKind.True:
                            case JsonValueKind.False:
                                value = cell.GetBoolean();
                                break;
                            case JsonValueKind.Null:
                                value = null;
                                break;
                            default:
                                // For other kinds (e.g., objects), store the raw JSON text
                                value = cell.GetRawText();
                                break;
                        }

                        // Populate the cell; PutValue handles nulls and various data types
                        sheet.Cells[rowIndex, colIndex].PutValue(value);
                        colIndex++;
                    }

                    rowIndex++;
                }
            }

            // Save the workbook to an Excel file
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
