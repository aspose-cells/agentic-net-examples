// Title: Batch generate timestamped Excel reports from JSON using Aspose.Cells smart markers (C#)
// Description: Scans a directory for JSON files, loads an Excel template with smart markers, binds each JSON string via WorkbookDesigner.SetJsonDataSource, processes the markers, and saves a workbook per file with a unique timestamped name.
// Keywords: Aspose.Cells | C# | smart markers | JSON data source | WorkbookDesigner | batch Excel generation | timestamped filenames | automated reporting | template merging
// Common Searches: Aspose.Cells set JSON data source C# example | How to use smart markers with JSON in Aspose.Cells | Create timestamped Excel files with Aspose.Cells | Batch process multiple JSON files into Excel reports | WorkbookDesigner SetJsonDataSource usage
// Developer Intent: Automatically merge each JSON file into a smart‑marker template and save the result as a uniquely timestamped Excel workbook.
// Use Cases: Daily sales dashboards: inject JSON sales data into a pre‑designed template and archive each report with a date‑time suffix. | Personalized invoices: convert per‑customer JSON files into formatted invoices, naming each file with a timestamp for audit trails. | Configuration dashboards: transform a folder of JSON config files into Excel sheets for quick visual analysis, generating one workbook per file.
// AI Prompts: Show C# code that adds validation for missing smart markers before calling WorkbookDesigner.Process. | Demonstrate how to customize the timestamp format (e.g., yyyy-MM-dd_HH-mm-ss) in the output filename. | Explain how to apply conditional formatting to the generated workbook after processing the JSON data.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonWorkflow
{
    // Scans a directory for JSON files, loads an Excel template with smart markers, binds each JSON string via WorkbookDesigner.SetJsonDataSource, processes the markers, and saves a workbook per file with a unique timestamped name.
    public class JsonToExcelProcessor
    {
        /// <param name="templatePath">Full path to the Excel template containing smart markers.</param>
        /// <param name="jsonFolder">Folder that contains the JSON files to be processed.</param>
        /// <param name="outputFolder">Folder where the generated workbooks will be saved.</param>
        public static void ProcessJsonFiles(string templatePath, string jsonFolder, string outputFolder)
        {
            try
            {
                // Verify template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Ensure JSON source folder exists
                if (!Directory.Exists(jsonFolder))
                {
                    Console.WriteLine($"JSON source folder not found: {jsonFolder}");
                    return;
                }

                // Ensure output directory exists
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                // Get all JSON files in the source folder
                string[] jsonFiles = Directory.GetFiles(jsonFolder, "*.json", SearchOption.TopDirectoryOnly);

                if (jsonFiles.Length == 0)
                {
                    Console.WriteLine("No JSON files found to process.");
                    return;
                }

                foreach (string jsonFilePath in jsonFiles)
                {
                    try
                    {
                        // Read JSON content
                        string jsonContent = File.ReadAllText(jsonFilePath);

                        // Load the template workbook
                        Workbook workbook = new Workbook(templatePath);

                        // Initialize WorkbookDesigner with the loaded workbook
                        WorkbookDesigner designer = new WorkbookDesigner(workbook);

                        // Use a generic data source name; it can be any identifier you use in the template markers
                        const string dataSourceName = "DataSource";

                        // Set the JSON string as the data source for smart markers
                        designer.SetJsonDataSource(dataSourceName, jsonContent);

                        // Process all smart markers in the workbook
                        designer.Process();

                        // Build timestamped output filename
                        string jsonFileName = Path.GetFileNameWithoutExtension(jsonFilePath);
                        string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                        string outputFileName = $"{jsonFileName}_{timestamp}.xlsx";
                        string outputPath = Path.Combine(outputFolder, outputFileName);

                        // Save the populated workbook
                        workbook.Save(outputPath);

                        Console.WriteLine($"Generated: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{jsonFilePath}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Example usage
        public static void Main()
        {
            // Path to the Excel template that contains smart markers like &DataSource.Name, etc.
            string templatePath = @"C:\Templates\ReportTemplate.xlsx";

            // Folder containing JSON files to be merged into the template
            string jsonFolder = @"C:\Data\JsonFiles";

            // Folder where the generated workbooks will be stored
            string outputFolder = @"C:\Data\GeneratedReports";

            ProcessJsonFiles(templatePath, jsonFolder, outputFolder);

            Console.WriteLine("Processing completed.");
        }
    }
}
