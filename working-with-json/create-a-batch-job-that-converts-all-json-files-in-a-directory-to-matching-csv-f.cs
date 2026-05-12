using System;
using System.IO;
using Aspose.Cells;

class JsonToCsvBatchConverter
{
    static void Main()
    {
        // Directory containing JSON files
        string sourceDirectory = @"C:\Data\JsonFiles";
        // Directory where CSV files will be saved (can be the same as source)
        string outputDirectory = @"C:\Data\CsvFiles";

        // Verify source directory exists
        if (!Directory.Exists(sourceDirectory))
        {
            Console.WriteLine($"Source directory does not exist: {sourceDirectory}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputDirectory);

        // Get all JSON files in the source directory
        string[] jsonFiles = Directory.GetFiles(sourceDirectory, "*.json", SearchOption.TopDirectoryOnly);

        foreach (string jsonFilePath in jsonFiles)
        {
            try
            {
                // Load JSON into a workbook using JsonLoadOptions
                JsonLoadOptions loadOptions = new JsonLoadOptions();
                Workbook workbook = new Workbook(jsonFilePath, loadOptions);

                // Determine CSV file name (same base name, .csv extension)
                string csvFileName = Path.GetFileNameWithoutExtension(jsonFilePath) + ".csv";
                string csvFilePath = Path.Combine(outputDirectory, csvFileName);

                // Save workbook as CSV
                workbook.Save(csvFilePath, SaveFormat.Csv);

                Console.WriteLine($"Converted: {Path.GetFileName(jsonFilePath)} -> {csvFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting '{jsonFilePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}