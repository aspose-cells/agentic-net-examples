// Title: Batch convert JSON files to timestamped Excel workbooks using an Aspose.Cells template in C#
// AI Prompts: Write C# code that iterates over every *.json file in a directory, deserializes each file into a Dictionary<string,object>, loads a given .xlsx template with Aspose.Cells, writes keys to column A and values to column B starting at row 2, and saves the workbook with a filename that appends the current yyyyMMdd_HHmmss timestamp. | Create a C# routine that checks that the JSON input folder and template file exist, creates the output folder if needed, processes each JSON file while logging success or failure, and catches any exceptions thrown by Aspose.Cells during workbook generation. | Implement a helper method that generates a timestamp string in the format yyyyMMdd_HHmmss, combines it with the original JSON file name, and uses Aspose.Cells SaveFormat.Xlsx to write the populated workbook to the output directory.
// Common Searches: c# read multiple json files and fill an Aspose.Cells template with key value pairs | how to add timestamp to Excel file name when saving with Aspose.Cells | batch processing json to xlsx using Aspose.Cells smart markers | generate Excel reports from a folder of json files in .NET | Aspose.Cells populate worksheet from dictionary and save with unique name
// Tags: populate worksheet from dictionary using Aspose.Cells | save workbook with timestamped filename C# | batch json to xlsx conversion Aspose.Cells | validate input and template paths Aspose.Cells | exception handling during Excel generation Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

namespace JsonToExcelWorkflow
{
    // The program scans a configured directory for .json files, deserializes each into a Dictionary<string,object>, loads a predefined Excel template with Aspose.Cells, writes each key/value pair to columns A and B starting at row 2 of the first worksheet, and saves the populated workbook to an output folder using a filename that combines the original JSON name with a yyyyMMdd_HHmmss timestamp, while handling missing paths and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths configuration
            string jsonInputFolder = @"C:\Data\JsonFiles";                     // Folder containing JSON files
            string templatePath = @"C:\Data\Template\Template.xlsx";          // Path to the Excel template
            string outputFolder = @"C:\Data\Output";                          // Folder to save generated workbooks

            // Verify input folder
            if (!Directory.Exists(jsonInputFolder))
            {
                Console.WriteLine($"Input folder not found: {jsonInputFolder}");
                return;
            }

            // Verify template file
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all JSON files in the input folder
            string[] jsonFiles = Directory.GetFiles(jsonInputFolder, "*.json");

            foreach (string jsonFilePath in jsonFiles)
            {
                try
                {
                    // Read JSON content
                    string jsonContent = File.ReadAllText(jsonFilePath);

                    // Deserialize to a dictionary for flexible key/value handling
                    Dictionary<string, object> data = JsonSerializer.Deserialize<Dictionary<string, object>>(jsonContent);

                    // Load the Excel template
                    Workbook workbook = new Workbook(templatePath);
                    Worksheet sheet = workbook.Worksheets[0]; // Assuming data goes to the first worksheet

                    // Simple population logic:
                    // Write each key/value pair starting from row 2 (row 1 can be headers)
                    int startRow = 1; // zero‑based index (row 2 in Excel)
                    int keyColumn = 0; // Column A
                    int valueColumn = 1; // Column B

                    foreach (var kvp in data)
                    {
                        // Write the key
                        sheet.Cells[startRow, keyColumn].PutValue(kvp.Key);
                        // Write the value (convert to string to avoid type issues)
                        sheet.Cells[startRow, valueColumn].PutValue(kvp.Value?.ToString() ?? string.Empty);
                        startRow++;
                    }

                    // Create a timestamp for the filename
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string jsonFileNameWithoutExt = Path.GetFileNameWithoutExtension(jsonFilePath);
                    string outputFileName = $"{jsonFileNameWithoutExt}_{timestamp}.xlsx";
                    string outputPath = Path.Combine(outputFolder, outputFileName);

                    // Save the populated workbook
                    workbook.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Processed '{jsonFilePath}' -> '{outputPath}'");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{jsonFilePath}': {ex.Message}");
                }
            }

            Console.WriteLine("All files have been processed.");
        }
    }
}
