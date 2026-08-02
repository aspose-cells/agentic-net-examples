// Title: Batch generate timestamped Excel workbooks from JSON using Aspose.Cells smart markers (C#)
// Description: A C# console app that scans a folder for *.json files, loads an Excel template with smart markers, assigns each JSON string to a WorkbookDesigner data source, processes the markers, and saves the populated workbook with a filename that combines the original JSON name and a precise timestamp.
// Keywords: Aspose.Cells JSON data source | WorkbookDesigner smart markers C# | batch generate Excel from JSON | timestamped Excel filename | C# read JSON folder Aspose.Cells | automate Excel report generation | populate Excel template JSON
// Common Searches: Aspose.Cells set JSON data source with WorkbookDesigner | C# generate Excel files from a folder of JSON files | smart markers JSON example Aspose.Cells | batch process JSON to Excel with timestamped names | how to use Aspose.Cells WorkbookDesigner for JSON
// Developer Intent: Build a .NET utility that reads every JSON file in a directory, fills an Excel template via smart markers, and writes each result to an output folder using a unique timestamped filename.
// Use Cases: Automate daily report creation by converting exported JSON data into formatted Excel workbooks. | Generate personalized invoices or statements from JSON order files using a single smart‑marker template. | Archive processed datasets with version‑controlled, timestamped Excel files for audit and traceability.
// AI Prompts: Write C# code that iterates over JSON files in a directory, uses Aspose.Cells WorkbookDesigner to set a JSON data source named 'DataSource', processes smart markers, and saves each workbook with a timestamped filename. | Explain how to configure smart‑marker prefixes in an Excel template to match the JSON data source name used by Aspose.Cells Designer. | Provide best‑practice error handling and logging for batch processing JSON files into Excel workbooks with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonWorkflow
{
    // A C# console app that scans a folder for *.json files, loads an Excel template with smart markers, assigns each JSON string to a WorkbookDesigner data source, processes the markers, and saves the populated workbook with a filename that combines the original JSON name and a precise timestamp.
    class Program
    {
        static void Main()
        {
            // Path to the folder containing JSON files
            string jsonFolder = @"C:\Data\JsonFiles";

            // Path to the Excel template that contains smart markers
            string templatePath = @"C:\Data\Template.xlsx";

            // Output folder for the generated workbooks
            string outputFolder = @"C:\Data\GeneratedWorkbooks";

            try
            {
                // Verify that the JSON folder exists
                if (!Directory.Exists(jsonFolder))
                {
                    Console.WriteLine($"JSON folder not found: {jsonFolder}");
                    return;
                }

                // Verify that the template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Get all JSON files in the source folder
                string[] jsonFiles = Directory.GetFiles(jsonFolder, "*.json");

                foreach (string jsonFile in jsonFiles)
                {
                    try
                    {
                        // Read JSON content
                        string jsonContent = File.ReadAllText(jsonFile);

                        // Load the template workbook
                        Workbook workbook = new Workbook(templatePath);

                        // Initialize WorkbookDesigner with the loaded workbook
                        WorkbookDesigner designer = new WorkbookDesigner(workbook);

                        // Set the JSON data source; the name "DataSource" must match the smart marker prefix in the template
                        designer.SetJsonDataSource("DataSource", jsonContent);

                        // Process the smart markers and populate the workbook
                        designer.Process();

                        // Build a timestamped file name: OriginalJsonFileName_yyyyMMdd_HHmmssfff.xlsx
                        string jsonFileName = Path.GetFileNameWithoutExtension(jsonFile);
                        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
                        string outputFileName = $"{jsonFileName}_{timestamp}.xlsx";
                        string outputPath = Path.Combine(outputFolder, outputFileName);

                        // Save the populated workbook
                        workbook.Save(outputPath, SaveFormat.Xlsx);

                        Console.WriteLine($"Generated workbook: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{jsonFile}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
