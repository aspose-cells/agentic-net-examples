// Title: Batch fill Aspose.Cells smart markers in multiple Excel templates with a JSON data source using C#
// AI Prompts: Generate C# code that scans a folder for .xlsx files, loads each workbook into a WorkbookDesigner, assigns a JSON data source named 'DataSource' to the smart markers, processes them, and saves the results to a specified output directory. | Modify the batch example to read an XML data source instead of JSON while keeping the same folder traversal and smart‑marker processing logic. | Add comprehensive logging to the batch routine that records the start time, each file processed, success or error status, and a summary report at the end.
// Common Searches: aspocells batch processing smart markers C# folder of templates | how to apply the same JSON data to many Excel files with WorkbookDesigner | C# loop through directory and populate smart markers in each workbook | save processed Excel workbooks to a different folder using Aspose.Cells | error handling for batch smart marker conversion in C#
// Tags: bulk smart marker handling with WorkbookDesigner | populate Excel templates from JSON using Aspose.Cells | iterate over .xlsx files in C# and apply data source | save processed workbooks to separate output directory | exception handling for bulk smart marker execution

using System;
using System.IO;
using Aspose.Cells;

namespace BatchSmartMarkerProcessing
{
    // The program enumerates all .xlsx files in a source folder, loads each workbook, sets a JSON data source for the smart markers via WorkbookDesigner, processes the markers, and writes the populated workbooks to a target folder, with basic error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing template workbooks with smart markers
            string inputFolder = @"C:\Templates";

            // Folder where processed workbooks will be saved
            string outputFolder = @"C:\Processed";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify that the input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Sample JSON data source that will be applied to every workbook
            string jsonData = @"{""Name"":""John Doe"",""Age"":30,""City"":""New York""}";

            try
            {
                // Process each Excel file in the input folder
                foreach (string templatePath in Directory.GetFiles(inputFolder, "*.xlsx"))
                {
                    // Ensure the template file still exists before loading
                    if (!File.Exists(templatePath))
                    {
                        Console.WriteLine($"File not found (skipped): {templatePath}");
                        continue;
                    }

                    try
                    {
                        // Load the template workbook
                        Workbook workbook = new Workbook(templatePath);

                        // Initialize the WorkbookDesigner with the loaded workbook
                        WorkbookDesigner designer = new WorkbookDesigner(workbook);

                        // Set the JSON data source (the name "DataSource" must match the smart marker prefix in the templates)
                        designer.SetJsonDataSource("DataSource", jsonData);

                        // Process all smart markers in the workbook
                        designer.Process();

                        // Determine the output file path (same file name, different folder)
                        string outputPath = Path.Combine(outputFolder, Path.GetFileName(templatePath));

                        // Save the processed workbook
                        workbook.Save(outputPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{templatePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
