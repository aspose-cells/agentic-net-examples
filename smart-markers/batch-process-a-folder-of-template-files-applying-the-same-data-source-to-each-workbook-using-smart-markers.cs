using System;
using System.IO;
using Aspose.Cells;

namespace BatchSmartMarkerProcessing
{
    class Program
    {
        // Sample JSON data that will be applied to every template workbook
        private const string JsonData = @"{
            ""Name"": ""John Doe"",
            ""Age"": 30,
            ""City"": ""New York""
        }";

        static void Main(string[] args)
        {
            // Folder that contains the template files (with smart markers)
            string inputFolder = @"C:\Templates";

            // Folder where processed files will be saved
            string outputFolder = @"C:\Processed";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure the output folder exists
            Directory.CreateDirectory(outputFolder);

            string[] templateFiles;
            try
            {
                // Get all Excel files in the input folder
                templateFiles = Directory.GetFiles(inputFolder, "*.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to enumerate files in '{inputFolder}': {ex.Message}");
                return;
            }

            foreach (string templatePath in templateFiles)
            {
                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(templatePath))
                    {
                        Console.WriteLine($"File not found (skipped): {templatePath}");
                        continue;
                    }

                    // Load the template workbook
                    Workbook workbook = new Workbook(templatePath);

                    // Initialize the designer with the loaded workbook
                    WorkbookDesigner designer = new WorkbookDesigner
                    {
                        Workbook = workbook
                    };

                    // Set the same JSON data source for all workbooks
                    designer.SetJsonDataSource("Data", JsonData);

                    // Process all smart markers in the workbook
                    designer.Process();

                    // Build the output file name (same as input but placed in output folder)
                    string fileName = Path.GetFileName(templatePath);
                    string outputPath = Path.Combine(outputFolder, fileName);

                    // Save the processed workbook
                    designer.Workbook.Save(outputPath);

                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{templatePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}