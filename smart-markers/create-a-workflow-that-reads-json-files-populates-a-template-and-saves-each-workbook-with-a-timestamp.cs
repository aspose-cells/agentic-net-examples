using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonWorkflow
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel template that contains smart markers
            string templatePath = "Template.xlsx";

            // Directory containing JSON files to be processed
            string jsonDirectory = "JsonData";

            // Directory where the generated workbooks will be saved
            string outputDirectory = "Output";

            try
            {
                // Verify template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Verify JSON directory exists
                if (!Directory.Exists(jsonDirectory))
                {
                    Console.WriteLine($"JSON directory not found: {jsonDirectory}");
                    return;
                }

                // Ensure the output directory exists
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // Get all JSON files in the specified directory
                string[] jsonFiles = Directory.GetFiles(jsonDirectory, "*.json");

                foreach (string jsonFilePath in jsonFiles)
                {
                    try
                    {
                        // Load the JSON content as a string
                        string jsonContent = File.ReadAllText(jsonFilePath);

                        // Load the template workbook
                        Workbook workbook = new Workbook(templatePath);

                        // Create a WorkbookDesigner and assign the loaded workbook
                        WorkbookDesigner designer = new WorkbookDesigner
                        {
                            Workbook = workbook
                        };

                        // Use the file name (without extension) as the data source name
                        string dataSourceName = Path.GetFileNameWithoutExtension(jsonFilePath);

                        // Set the JSON data source for smart markers
                        designer.SetJsonDataSource(dataSourceName, jsonContent);

                        // Process the smart markers to populate the workbook
                        designer.Process();

                        // Build a timestamp string for the output file name
                        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                        // Construct the output file name
                        string outputFileName = $"{Path.GetFileNameWithoutExtension(jsonFilePath)}_{timestamp}.xlsx";
                        string outputPath = Path.Combine(outputDirectory, outputFileName);

                        // Save the populated workbook
                        workbook.Save(outputPath, SaveFormat.Xlsx);

                        Console.WriteLine($"Processed '{jsonFilePath}' -> '{outputPath}'");
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
    }
}