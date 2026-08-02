// Title: C# Batch Processing of Excel Templates with JSON via Aspose.Cells Smart Markers
// Description: A C# console application that scans a directory of Excel templates, pairs each file with a JSON file of the same base name, loads the workbook, sets the JSON as a smart‑marker data source, processes all smart markers, and writes the populated workbook to an output folder. Supports .xlsx and .xls files, validates folder existence, and provides robust error handling.
// Keywords: Aspose.Cells | Smart Markers | C# batch processing | Excel template automation | JSON data source | WorkbookDesigner | bulk Excel generation | template to JSON mapping | automated report creation | invoice generation C#
// Common Searches: batch process Excel templates with JSON using Aspose.Cells C# | C# loop through folder of Excel files and apply smart markers | Aspose.Cells WorkbookDesigner bulk processing example | automate Excel generation from JSON files C# | process multiple workbooks with smart markers Aspose
// Developer Intent: Create a batch job that loads each Excel template, applies its matching JSON data via smart markers, and saves the result.
// Use Cases: Generate personalized invoices by pairing invoice templates with customer‑specific JSON files. | Produce monthly performance reports where each period’s template is filled with data from a correspondingly named JSON file. | Automate issuance of employee certificates by matching a certificate template with JSON files containing individual employee details.
// AI Prompts: Write a reusable C# method that accepts template, JSON, and output directories and processes all matching file pairs using Aspose.Cells WorkbookDesigner. | Suggest best‑practice error handling for batch smart‑marker processing when some templates or JSON files are missing. | Show how to make the smart‑marker data source name configurable instead of hard‑coded "Data". | Provide a PowerShell script that invokes the compiled batch processor for scheduled execution. | Explain how to extend the code to handle nested JSON objects with hierarchical smart markers.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchSmartMarkerProcessor
{
    // A C# console application that scans a directory of Excel templates, pairs each file with a JSON file of the same base name, loads the workbook, sets the JSON as a smart‑marker data source, processes all smart markers, and writes the populated workbook to an output folder. Supports .xlsx and .xls files, validates folder existence, and provides robust error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output directories (adjust as needed)
            string templatesFolder = @"C:\Templates";          // Folder containing Excel templates
            string jsonDataFolder = @"C:\JsonData";            // Folder containing JSON files
            string outputFolder = @"C:\ProcessedOutputs";

            try
            {
                // Verify that the required folders exist
                if (!Directory.Exists(templatesFolder))
                {
                    Console.WriteLine($"Templates folder not found: {templatesFolder}");
                    return;
                }

                if (!Directory.Exists(jsonDataFolder))
                {
                    Console.WriteLine($"JSON data folder not found: {jsonDataFolder}");
                    return;
                }

                // Ensure output directory exists
                Directory.CreateDirectory(outputFolder);

                // Get all Excel template files (supports .xlsx and .xls)
                string[] templateFiles = Directory.GetFiles(templatesFolder, "*.*", SearchOption.TopDirectoryOnly);
                foreach (string templatePath in templateFiles)
                {
                    string extension = Path.GetExtension(templatePath).ToLowerInvariant();
                    if (extension != ".xlsx" && extension != ".xls")
                        continue; // Skip non-Excel files

                    // Determine corresponding JSON file (same base name, .json extension)
                    string baseName = Path.GetFileNameWithoutExtension(templatePath);
                    string jsonPath = Path.Combine(jsonDataFolder, baseName + ".json");

                    if (!File.Exists(jsonPath))
                    {
                        Console.WriteLine($"JSON data file not found for template '{baseName}'. Skipping.");
                        continue;
                    }

                    // Verify the template file still exists before loading
                    if (!File.Exists(templatePath))
                    {
                        Console.WriteLine($"Template file not found: {templatePath}. Skipping.");
                        continue;
                    }

                    try
                    {
                        // Load the Excel template
                        Workbook workbook = new Workbook(templatePath);

                        // Read JSON content
                        string jsonData = File.ReadAllText(jsonPath);

                        // Set up WorkbookDesigner with the loaded workbook
                        WorkbookDesigner designer = new WorkbookDesigner(workbook);

                        // Use a generic data source name; smart markers in the template should reference this name
                        const string dataSourceName = "Data";
                        designer.SetJsonDataSource(dataSourceName, jsonData);

                        // Process all smart markers in the workbook
                        designer.Process();

                        // Save the processed workbook to the output folder
                        string outputPath = Path.Combine(outputFolder, baseName + "_Processed.xlsx");
                        workbook.Save(outputPath);

                        Console.WriteLine($"Processed '{baseName}' and saved to '{outputPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing template '{baseName}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
