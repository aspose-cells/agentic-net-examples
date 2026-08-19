// Title: C# Batch Processing of Excel Templates with Smart Markers and JSON using Aspose.Cells
// Description: A console app that scans a folder for .xlsx/.xls templates, pairs each with a same‑named .json file, loads the workbook with Aspose.Cells, assigns the JSON via WorkbookDesigner, processes all smart markers, and saves the populated files to an output directory with robust error handling.
// Keywords: Aspose.Cells | C# | smart markers | batch Excel processing | JSON data source | WorkbookDesigner | .NET Excel automation | template population | invoice generation | certificate automation
// Common Searches: Aspose.Cells batch smart markers example | C# process multiple Excel templates with JSON | WorkbookDesigner SetJsonDataSource loop folder | automate Excel report generation using Aspose | smart marker batch processing C#
// Developer Intent: Create a C# utility that automatically fills a collection of Excel templates with corresponding JSON data via Aspose.Cells smart markers.
// Use Cases: Generate a set of invoices by matching each invoice template with a customer‑specific JSON file. | Produce personalized certificates for event participants using template workbooks and individual JSON records. | Automate monthly financial reports by populating multiple report templates with JSON extracts from a database.
// AI Prompts: Write a function that logs batch processing results to a CSV file instead of the console. | Suggest enhancements to validate JSON schema before assigning it to WorkbookDesigner. | Show how to configure the processor to use a custom data‑source name derived from a JSON property.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchSmartMarkerProcessor
{
    // A console app that scans a folder for .xlsx/.xls templates, pairs each with a same‑named .json file, loads the workbook with Aspose.Cells, assigns the JSON via WorkbookDesigner, processes all smart markers, and saves the populated files to an output directory with robust error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing Excel templates
            string templatesFolder = @"C:\Templates";

            // Folder where processed workbooks will be saved
            string outputFolder = @"C:\Processed";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify that the templates folder exists
            if (!Directory.Exists(templatesFolder))
            {
                Console.WriteLine($"Templates folder not found: {templatesFolder}");
                Console.WriteLine("Please create the folder and add template files before running the program.");
                return;
            }

            string[] excelFiles;
            try
            {
                // Get all Excel files in the templates folder (supports .xlsx and .xls)
                excelFiles = Directory.GetFiles(templatesFolder, "*.*", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error accessing templates folder: {ex.Message}");
                return;
            }

            foreach (string excelPath in excelFiles)
            {
                string extension = Path.GetExtension(excelPath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls")
                    continue; // Skip non‑Excel files

                // Determine corresponding JSON file (same base name, .json extension)
                string jsonPath = Path.ChangeExtension(excelPath, ".json");
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"JSON data not found for template '{Path.GetFileName(excelPath)}'. Skipping.");
                    continue;
                }

                try
                {
                    // Load the Excel template
                    Workbook workbook = new Workbook(excelPath);

                    // Read JSON content
                    string jsonData = File.ReadAllText(jsonPath);

                    // Create a WorkbookDesigner and assign the workbook
                    WorkbookDesigner designer = new WorkbookDesigner
                    {
                        Workbook = workbook
                    };

                    // Use the file name (without extension) as the data source name
                    string dataSourceName = Path.GetFileNameWithoutExtension(jsonPath);
                    designer.SetJsonDataSource(dataSourceName, jsonData);

                    // Process all smart markers in the workbook
                    designer.Process();

                    // Save the processed workbook to the output folder
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(excelPath));
                    workbook.Save(outputPath);

                    Console.WriteLine($"Processed '{Path.GetFileName(excelPath)}' successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(excelPath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
