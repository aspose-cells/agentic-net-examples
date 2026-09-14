// Title: Batch generate timestamped Excel workbooks from JSON files using Aspose.Cells smart markers in C#
// AI Prompts: Create a C# console program that iterates over every .json file in a directory, loads each file as a JSON data source for a WorkbookDesigner, processes smart markers in a supplied Excel template, and saves the populated workbook as an .xlsx file whose name combines the original JSON filename with the current date‑time. | Extend the batch workflow so the desired output format (e.g., XLSX or PDF) can be selected at runtime while preserving the same date‑time based naming convention for the generated files. | Add robust error handling that catches JSON parsing or processing exceptions, writes the offending file path and exception details to a log, and then continues processing the remaining files.
// Common Searches: how to use Aspose.Cells WorkbookDesigner to populate an Excel template from multiple JSON files in C# | c# batch process json files into excel workbooks with smart markers and timestamped filenames | aspocells setjsondatasource example for generating reports from json data | save Aspose.Cells workbook with dynamic datetime in filename c# | convert populated Excel template to PDF using Aspose.Cells while keeping timestamped output name
// Tags: Aspose.Cells WorkbookDesigner JSON data source | batch Excel generation with JSON data C# | date-time based workbook naming Aspose.Cells | configurable output format Aspose.Cells | JSON parsing error logging Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonWorkflow
{
    // The example scans a folder for .json files, loads each file as a JSON data source for a WorkbookDesigner, processes smart markers in a predefined Excel template, and saves the resulting workbook to an output directory. Each generated file is named using the source JSON filename plus a current timestamp, ensuring unique, date‑time based filenames.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the folder containing JSON files
            string jsonFolder = @"C:\Data\JsonFiles";

            // Path to the Excel template that contains smart markers
            string templatePath = @"C:\Data\Template.xlsx";

            // Output folder for the generated workbooks
            string outputFolder = @"C:\Data\GeneratedWorkbooks";

            // Verify required directories and files exist
            if (!Directory.Exists(jsonFolder))
            {
                Console.WriteLine($"JSON folder not found: {jsonFolder}");
                return;
            }

            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all JSON files in the source folder
            string[] jsonFiles = Directory.GetFiles(jsonFolder, "*.json");

            foreach (string jsonFilePath in jsonFiles)
            {
                try
                {
                    // Read JSON content
                    string jsonContent = File.ReadAllText(jsonFilePath);

                    // Load the template workbook
                    Workbook workbook = new Workbook(templatePath);

                    // Create a WorkbookDesigner and assign the workbook
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);

                    // Use the file name (without extension) as the data source name
                    string dataSourceName = Path.GetFileNameWithoutExtension(jsonFilePath);

                    // Set the JSON data source for smart markers
                    designer.SetJsonDataSource(dataSourceName, jsonContent);

                    // Process the smart markers to populate data
                    designer.Process();

                    // Build output file name with timestamp
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string outputFileName = $"{dataSourceName}_{timestamp}.xlsx";
                    string outputPath = Path.Combine(outputFolder, outputFileName);

                    // Save the populated workbook
                    workbook.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Generated workbook: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{jsonFilePath}': {ex.Message}");
                }
            }
        }
    }
}
