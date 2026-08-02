// Title: Batch process Excel templates with a shared JSON data source using Aspose.Cells smart markers in C#
// Description: C# utility that scans a directory for .xlsx templates, loads each workbook into a WorkbookDesigner, assigns a single JSON data source, processes all smart markers, and saves the populated files to a target folder while handling missing files and runtime errors.
// Keywords: Aspose.Cells batch processing | smart markers C# | WorkbookDesigner SetJsonDataSource | populate multiple Excel templates | folder‑wide Excel automation | .NET Excel report generation | JSON data source Excel
// Common Searches: Aspose.Cells process all Excel files in a folder | apply same JSON to multiple workbooks using smart markers | C# loop through .xlsx templates and fill data | batch generate reports with Aspose.Cells WorkbookDesigner | automate smart marker population for many files
// Developer Intent: Use one JSON payload to fill smart markers in every Excel template within a folder and write the completed workbooks to a separate output directory.
// Use Cases: Produce a batch of personalized invoices by applying a common customer JSON to each invoice template. | Refresh a set of regional sales dashboards with identical sales figures stored in JSON. | Execute mass mail‑merge by populating smart markers across dozens of letter templates with shared recipient data.
// AI Prompts: Add comprehensive logging (including file name and processing time) to the batch smart‑marker loop in C#. | Rewrite the example to consume an XML data source instead of JSON while keeping the folder‑wide workflow. | Show how to rename each output file using a value extracted from a smart marker inside the source workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchSmartMarkerProcessing
{
    // C# utility that scans a directory for .xlsx templates, loads each workbook into a WorkbookDesigner, assigns a single JSON data source, processes all smart markers, and saves the populated files to a target folder while handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing template workbooks with smart markers
            string inputFolder = @"C:\Templates";
            // Folder where processed workbooks will be saved
            string outputFolder = @"C:\Processed";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure the output folder exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Sample JSON data source that will be applied to every workbook
            string jsonData = @"{
                ""Name"": ""John Doe"",
                ""Age"": 30,
                ""City"": ""New York"",
                ""Products"": [
                    { ""ProductID"": 1, ""ProductName"": ""Apple"", ""Price"": 1.2 },
                    { ""ProductID"": 2, ""ProductName"": ""Banana"", ""Price"": 0.8 }
                ]
            }";

            // Get all Excel template files (you can adjust the search pattern as needed)
            string[] templateFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string templatePath in templateFiles)
            {
                try
                {
                    // Verify the template file exists before loading
                    if (!File.Exists(templatePath))
                    {
                        Console.WriteLine($"Template file not found: {templatePath}");
                        continue;
                    }

                    // Load the template workbook
                    Workbook workbook = new Workbook(templatePath);

                    // Initialize the designer with the loaded workbook
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);

                    // Set the same JSON data source for all workbooks
                    designer.SetJsonDataSource("DataSource", jsonData);

                    // Process smart markers in the workbook
                    designer.Process();

                    // Determine the output file path (same file name, different folder)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(templatePath));

                    // Save the processed workbook
                    workbook.Save(outputPath);

                    Console.WriteLine($"Processed: {Path.GetFileName(templatePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{templatePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed. Processed files are saved to: " + outputFolder);
        }
    }
}
