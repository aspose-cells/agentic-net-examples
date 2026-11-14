// Title: C# console batch job to fill Excel (.xlsx) templates with matching JSON files using Aspose.Cells
// Common Searches: how to batch fill Excel templates from JSON files using Aspose.Cells in C# | C# Aspose.Cells example for processing a folder of .xlsx templates with corresponding .json data | populate multiple Excel workbooks with JSON data in a .NET console application | automate filling cell addresses from a JSON dictionary with Aspose.Cells | save populated Excel files to an output directory using Aspose.Cells C#
// Tags: batch fill Excel templates Aspose.Cells C# | populate .xlsx cells from JSON dictionary Aspose.Cells | process folder of Excel templates with matching JSON files | convert JsonElement to .NET primitive for Excel cell values | save filled workbooks to output folder Aspose.Cells | load JSON into Aspose.Cells cell mapping

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace ExcelBatchProcessor
{
    // The C# console program scans a templates folder for .xlsx files, loads each workbook with Aspose.Cells, reads a same‑named .json file into a dictionary of cell addresses and values, writes those values into the worksheet cells, and saves the populated workbook to a designated output directory.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing Excel templates (*.xlsx) and matching JSON files (*.json)
            string templatesFolder = @"C:\Templates";
            // Folder where the populated workbooks will be saved
            string outputFolder = @"C:\Output";

            // Verify that the templates folder exists
            if (!Directory.Exists(templatesFolder))
            {
                Console.WriteLine($"Templates folder not found: {templatesFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each Excel template in the folder
            foreach (string templatePath in Directory.GetFiles(templatesFolder, "*.xlsx"))
            {
                try
                {
                    // Load the Excel template
                    Workbook workbook = new Workbook(templatePath);
                    Worksheet worksheet = workbook.Worksheets[0];
                    Cells cells = worksheet.Cells;

                    // Determine the corresponding JSON file (same name, .json extension)
                    string jsonPath = Path.ChangeExtension(templatePath, ".json");
                    if (!File.Exists(jsonPath))
                    {
                        Console.WriteLine($"JSON data not found for template: {Path.GetFileName(templatePath)}. Skipping.");
                        continue;
                    }

                    // Load JSON data into a dictionary (cell address => value)
                    string jsonContent = File.ReadAllText(jsonPath);
                    var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                    var rawData = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(jsonContent, options);
                    if (rawData == null)
                    {
                        Console.WriteLine($"Failed to parse JSON for template: {Path.GetFileName(templatePath)}. Skipping.");
                        continue;
                    }

                    // Populate cells based on JSON mapping
                    foreach (KeyValuePair<string, JsonElement> entry in rawData)
                    {
                        object value = ConvertJsonElement(entry.Value);
                        cells[entry.Key].PutValue(value);
                    }

                    // Build output file name
                    string outputFileName = Path.GetFileNameWithoutExtension(templatePath) + "_filled.xlsx";
                    string outputPath = Path.Combine(outputFolder, outputFileName);

                    // Save the populated workbook
                    workbook.Save(outputPath);
                    Console.WriteLine($"Processed and saved: {outputFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing template '{Path.GetFileName(templatePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }

        // Convert JsonElement to a .NET primitive that Aspose.Cells can handle
        private static object ConvertJsonElement(JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.String => element.GetString(),
                JsonValueKind.Number => element.TryGetInt64(out long l) ? (object)l : element.GetDouble(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null => null,
                _ => element.GetRawText()
            };
        }
    }
}
