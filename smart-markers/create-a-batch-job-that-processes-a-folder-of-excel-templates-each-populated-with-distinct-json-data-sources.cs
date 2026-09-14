// Title: C# batch job to populate Excel .xlsx templates with JSON data using Aspose.Cells smart markers
// AI Prompts: Generate C# code that scans a directory for .xlsx template files, loads each workbook with Aspose.Cells, assigns a same‑named JSON file as a smart‑marker data source via WorkbookDesigner, processes the markers, and writes the result to an output folder. | Write a C# console application that iterates over Excel templates, reads corresponding JSON files, uses WorkbookDesigner.SetJsonDataSource to bind the data, calls Process, and saves the populated workbooks with a "_Processed" suffix.
// Common Searches: how to use Aspose.Cells WorkbookDesigner to apply JSON data to multiple Excel templates in C# | C# script for batch processing of .xlsx files with smart markers and JSON sources | automate population of Excel smart markers from JSON files using Aspose.Cells library
// Tags: batch populate Excel smart markers via Aspose.Cells | WorkbookDesigner JSON data source for .xlsx templates | C# loop over Excel files with smart markers | automated smart marker processing from JSON files | save processed workbooks with suffix in C#

using System;
using System.IO;
using Aspose.Cells;

namespace BatchSmartMarkerProcessor
{
    // Scans a folder of .xlsx templates, matches each with a same‑named .json file, loads the workbook, sets the JSON as a data source for smart markers via WorkbookDesigner, processes the markers, and saves the populated workbook to an output directory with a "_Processed" suffix.
    class Program
    {
        static void Main()
        {
            // Folder containing Excel template files (must contain smart markers)
            string templatesFolder = @"C:\Templates";

            // Folder containing JSON data files. Each JSON file should have the same base name as its template.
            string jsonFolder = @"C:\JsonData";

            // Folder where processed workbooks will be saved
            string outputFolder = @"C:\Processed";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify that the templates folder exists
            if (!Directory.Exists(templatesFolder))
            {
                Console.WriteLine($"Templates folder not found: '{templatesFolder}'.");
                return;
            }

            // Iterate over all .xlsx files in the templates folder
            foreach (string templatePath in Directory.GetFiles(templatesFolder, "*.xlsx"))
            {
                try
                {
                    // Verify that the template file actually exists (safety check)
                    if (!File.Exists(templatePath))
                    {
                        Console.WriteLine($"Template file not found: '{templatePath}'. Skipping.");
                        continue;
                    }

                    // Determine the base file name (without extension)
                    string baseName = Path.GetFileNameWithoutExtension(templatePath);

                    // Build the expected JSON file path
                    string jsonPath = Path.Combine(jsonFolder, baseName + ".json");

                    // Verify that the JSON file exists; if not, skip this template
                    if (!File.Exists(jsonPath))
                    {
                        Console.WriteLine($"JSON data not found for template '{baseName}'. Skipping.");
                        continue;
                    }

                    // Load the template workbook
                    Workbook workbook = new Workbook(templatePath);

                    // Create a WorkbookDesigner and associate it with the loaded workbook
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);

                    // Read the JSON content
                    string jsonData = File.ReadAllText(jsonPath);

                    // Set the JSON data source. The name "Data" can be referenced in smart markers (e.g., &=$Data.Name)
                    designer.SetJsonDataSource("Data", jsonData);

                    // Process all smart markers in the workbook
                    designer.Process();

                    // Build the output file path
                    string outputPath = Path.Combine(outputFolder, baseName + "_Processed.xlsx");

                    // Save the processed workbook
                    workbook.Save(outputPath);

                    Console.WriteLine($"Processed '{baseName}.xlsx' with data from '{baseName}.json' -> saved to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing template '{templatePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
