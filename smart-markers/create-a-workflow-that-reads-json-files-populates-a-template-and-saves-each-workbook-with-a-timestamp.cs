// Title: Generate timestamped Excel workbooks from JSON files using Aspose.Cells template copying in C#
// Common Searches: C# read multiple JSON files and fill an Excel template with Aspose.Cells | How to copy an Excel template and populate cells from a JSON dictionary using Aspose.Cells | Save Aspose.Cells workbook with a timestamped filename in .NET | Batch generate Excel files from JSON data with Aspose.Cells and include creation time | Aspose.Cells populate cells by address from JSON and export as .xlsx
// Tags: Aspose.Cells copy template workbook | populate cells from JSON dictionary Aspose.Cells | save workbook with timestamp filename C# | batch JSON to Excel conversion Aspose.Cells | deserialize JSON to cell address mapping .NET

using System;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsJsonWorkflow
{
    // The example iterates through all JSON files in a folder, deserializes each into a dictionary where keys are cell addresses, copies a predefined Excel template workbook, writes the values into the corresponding cells, appends a current timestamp to the original JSON filename, and saves the populated workbook as an .xlsx file in an output directory using Aspose.Cells for .NET.
    class Program
    {
        // Path to the folder containing JSON files
        private const string JsonFolderPath = @"C:\Data\JsonFiles";

        // Path to the Excel template file
        private const string TemplatePath = @"C:\Data\Template.xlsx";

        // Output folder for generated workbooks
        private const string OutputFolder = @"C:\Data\GeneratedWorkbooks";

        static void Main(string[] args)
        {
            try
            {
                // Ensure output directory exists
                Directory.CreateDirectory(OutputFolder);

                // Verify template file exists
                if (!File.Exists(TemplatePath))
                {
                    Console.WriteLine($"Template file not found: {TemplatePath}");
                    return;
                }

                // Load the template workbook once
                Workbook templateWorkbook = new Workbook(TemplatePath);

                // Verify JSON folder exists
                if (!Directory.Exists(JsonFolderPath))
                {
                    Console.WriteLine($"JSON folder not found: {JsonFolderPath}");
                    return;
                }

                // Get all JSON files in the specified folder
                string[] jsonFiles = Directory.GetFiles(JsonFolderPath, "*.json");
                if (jsonFiles.Length == 0)
                {
                    Console.WriteLine("No JSON files found.");
                    return;
                }

                foreach (string jsonFilePath in jsonFiles)
                {
                    try
                    {
                        // Read JSON content
                        string jsonContent = File.ReadAllText(jsonFilePath);

                        // Deserialize JSON into a dictionary where key = cell address (e.g., "A1")
                        var cellData = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(jsonContent);
                        if (cellData == null)
                        {
                            Console.WriteLine($"Failed to deserialize JSON: {jsonFilePath}");
                            continue;
                        }

                        // Create a new workbook and copy the template content
                        Workbook workbook = new Workbook();
                        workbook.Copy(templateWorkbook);

                        // Get the first worksheet (adjust if needed)
                        Worksheet sheet = workbook.Worksheets[0];

                        // Populate cells based on JSON data
                        foreach (KeyValuePair<string, JsonElement> kvp in cellData)
                        {
                            // Use the cell address from the JSON key
                            Cell cell = sheet.Cells[kvp.Key];

                            // Convert JsonElement to a .NET type that Aspose.Cells can handle
                            object value = kvp.Value.ValueKind switch
                            {
                                JsonValueKind.Number => kvp.Value.GetDouble(),
                                JsonValueKind.True => true,
                                JsonValueKind.False => false,
                                JsonValueKind.String => kvp.Value.GetString(),
                                _ => kvp.Value.ToString()
                            };

                            // Set the cell value; Aspose.Cells will handle type conversion
                            cell.PutValue(value);
                        }

                        // Build a timestamp string (e.g., 20231124_150530)
                        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                        // Create output file name: original JSON name + timestamp
                        string jsonFileName = Path.GetFileNameWithoutExtension(jsonFilePath);
                        string outputFileName = $"{jsonFileName}_{timestamp}.xlsx";
                        string outputPath = Path.Combine(OutputFolder, outputFileName);

                        // Save the populated workbook
                        workbook.Save(outputPath, SaveFormat.Xlsx);

                        Console.WriteLine($"Generated workbook: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{jsonFilePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
