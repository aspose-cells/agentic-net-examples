using System;
using System.IO;
using Aspose.Cells;

class BatchProcessor
{
    static void Main()
    {
        // Folder containing Excel templates and matching JSON files
        string templatesFolder = @"C:\Templates";
        // Folder where processed workbooks will be saved
        string outputFolder = @"C:\Processed";

        // Verify that the templates folder exists
        if (!Directory.Exists(templatesFolder))
        {
            Console.WriteLine($"Templates folder not found: {templatesFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all .xlsx template files in the folder
        string[] templateFiles = Directory.GetFiles(templatesFolder, "*.xlsx");

        foreach (string templatePath in templateFiles)
        {
            try
            {
                // Assume the JSON data source has the same base name as the template
                string jsonPath = Path.ChangeExtension(templatePath, ".json");

                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"JSON file not found for template '{Path.GetFileName(templatePath)}'. Skipping.");
                    continue;
                }

                // Verify the template file exists before loading
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}. Skipping.");
                    continue;
                }

                // Load the Excel template
                Workbook workbook = new Workbook(templatePath);

                // Initialize WorkbookDesigner with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Read JSON content
                string jsonData = File.ReadAllText(jsonPath);

                // Set JSON data source (using a fixed name "Data")
                designer.SetJsonDataSource("Data", jsonData);

                // Process smart markers; preserve unrecognized markers
                designer.Process(true);

                // Construct output file name
                string outputFileName = Path.GetFileNameWithoutExtension(templatePath) + "_filled.xlsx";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the processed workbook
                workbook.Save(outputPath);

                Console.WriteLine($"Processed '{Path.GetFileName(templatePath)}' -> '{outputFileName}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{Path.GetFileName(templatePath)}': {ex.Message}");
            }
        }
    }
}