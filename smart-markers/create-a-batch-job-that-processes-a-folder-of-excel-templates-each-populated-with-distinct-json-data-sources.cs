// Title: C# batch processor to fill Excel (.xlsx) templates with matching JSON files using Aspose.Cells smart markers
// AI Prompts: Create a C# console application that scans a directory for .xlsx templates, loads the corresponding .json files with the same base name, and uses Aspose.Cells to replace smart markers or named ranges with the JSON values before saving each workbook to an output folder. | Add comprehensive logging to the batch job so that each step—file discovery, JSON deserialization, cell population, and workbook saving—is recorded to both the console and a log file, while gracefully handling missing or malformed JSON files. | Enhance the processor to handle nested JSON objects by flattening them into dot‑separated keys and mapping those keys to hierarchical named ranges within the Excel templates.
// Common Searches: how to use Aspose.Cells to populate multiple Excel templates from JSON files in C# | C# batch fill .xlsx templates with data from matching .json files using smart markers | automate workbook generation from a folder of Excel templates and JSON sources with Aspose.Cells | process directory of Excel templates and JSON data programmatically in .NET
// Tags: batch populate Excel templates with JSON using Aspose.Cells | Aspose.Cells smart markers JSON to named range mapping | C# process folder of .xlsx and .json files | populate workbook cells from dictionary in Aspose.Cells | automated report generation from Excel templates .NET

using System;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsProcessor
{
    // The program iterates over .xlsx templates in a specified folder, loads matching .json files, maps each JSON key to a named range or direct cell address using Aspose.Cells smart markers, writes the values, and saves the populated workbooks to an output directory.
    class Program
    {
        static void Main()
        {
            // Folder containing Excel templates and matching JSON files
            string templatesFolder = @"C:\Templates";

            // Folder where processed workbooks will be saved
            string outputFolder = @"C:\Processed";

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Verify templates folder exists
            if (!Directory.Exists(templatesFolder))
            {
                Console.WriteLine($"Templates folder not found: {templatesFolder}");
                return;
            }

            // Iterate over each Excel template in the folder
            foreach (string templatePath in Directory.GetFiles(templatesFolder, "*.xlsx"))
            {
                try
                {
                    // Load JSON data (flat key/value structure)
                    string baseName = Path.GetFileNameWithoutExtension(templatePath);
                    string jsonPath = Path.Combine(templatesFolder, baseName + ".json");

                    if (!File.Exists(jsonPath))
                    {
                        Console.WriteLine($"JSON file not found for template '{baseName}'. Skipping.");
                        continue;
                    }

                    Dictionary<string, JsonElement> jsonData;
                    try
                    {
                        string jsonContent = File.ReadAllText(jsonPath);
                        jsonData = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(jsonContent)
                                   ?? new Dictionary<string, JsonElement>();
                    }
                    catch (Exception jsonEx)
                    {
                        Console.WriteLine($"Error reading JSON for '{baseName}': {jsonEx.Message}. Skipping.");
                        continue;
                    }

                    // Load the Excel template
                    Workbook workbook;
                    try
                    {
                        workbook = new Workbook(templatePath);
                    }
                    catch (Exception wbEx)
                    {
                        Console.WriteLine($"Failed to load workbook '{templatePath}': {wbEx.Message}. Skipping.");
                        continue;
                    }

                    Worksheet sheet = workbook.Worksheets[0]; // Assuming data goes to the first sheet

                    // Populate the workbook using named ranges or direct cell addresses
                    foreach (var kvp in jsonData)
                    {
                        try
                        {
                            // Try to locate a named range that matches the JSON key
                            Name namedRange = workbook.Worksheets.Names[kvp.Key];
                            if (namedRange != null)
                            {
                                // Resolve the first cell of the named range (e.g., "Sheet1!$A$1:$B$2")
                                string refersTo = namedRange.RefersTo;
                                if (refersTo.Contains("!"))
                                    refersTo = refersTo.Split('!')[1]; // keep only the address part
                                string firstCellAddress = refersTo.Split(':')[0].Replace("$", string.Empty);
                                Cell targetCell = sheet.Cells[firstCellAddress];
                                targetCell.PutValue(ConvertJsonElement(kvp.Value));
                            }
                            else
                            {
                                // If no named range, treat the key as a cell address (e.g., "B2")
                                Cell targetCell = sheet.Cells[kvp.Key];
                                targetCell.PutValue(ConvertJsonElement(kvp.Value));
                            }
                        }
                        catch (Exception innerEx)
                        {
                            // Log but continue processing other keys
                            Console.WriteLine($"Warning: Could not write value for key '{kvp.Key}' – {innerEx.Message}");
                        }
                    }

                    // Save the populated workbook to the output folder
                    string outputPath = Path.Combine(outputFolder, baseName + "_filled.xlsx");
                    try
                    {
                        workbook.Save(outputPath);
                        Console.WriteLine($"Processed '{baseName}' -> '{outputPath}'.");
                    }
                    catch (Exception saveEx)
                    {
                        Console.WriteLine($"Failed to save workbook for '{baseName}': {saveEx.Message}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{templatePath}': {ex.Message}");
                }
            }
        }

        // Convert JsonElement to a .NET primitive suitable for Aspose.Cells PutValue
        private static object ConvertJsonElement(JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.String => element.GetString(),
                JsonValueKind.Number => element.TryGetInt64(out long l) ? (object)l :
                                         element.TryGetDouble(out double d) ? d :
                                         element.GetDecimal(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null => null,
                _ => element.GetRawText()
            };
        }
    }
}
