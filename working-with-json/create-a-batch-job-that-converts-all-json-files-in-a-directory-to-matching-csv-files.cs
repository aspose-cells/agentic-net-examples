using System;
using System.IO;
using Aspose.Cells;

namespace JsonToCsvBatch
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Directory containing JSON files
                string sourceDirectory = @"C:\Data\JsonFiles";
                // Directory where CSV files will be saved
                string outputDirectory = @"C:\Data\CsvFiles";

                // Verify source directory exists
                if (!Directory.Exists(sourceDirectory))
                {
                    Console.WriteLine($"Source directory not found: {sourceDirectory}");
                    return;
                }

                // Ensure output directory exists
                Directory.CreateDirectory(outputDirectory);

                // Get all .json files in the source directory
                string[] jsonFiles = Directory.GetFiles(sourceDirectory, "*.json", SearchOption.TopDirectoryOnly);

                foreach (string jsonFilePath in jsonFiles)
                {
                    try
                    {
                        // Verify the JSON file still exists
                        if (!File.Exists(jsonFilePath))
                        {
                            Console.WriteLine($"File not found (skipped): {jsonFilePath}");
                            continue;
                        }

                        // Load JSON file into a workbook using JsonLoadOptions
                        JsonLoadOptions loadOptions = new JsonLoadOptions();
                        Workbook workbook = new Workbook(jsonFilePath, loadOptions);

                        // Determine output CSV file path (same name, .csv extension)
                        string csvFileName = Path.GetFileNameWithoutExtension(jsonFilePath) + ".csv";
                        string csvFilePath = Path.Combine(outputDirectory, csvFileName);

                        // Save workbook as CSV
                        workbook.Save(csvFilePath, SaveFormat.Csv);

                        Console.WriteLine($"Converted: {jsonFilePath} -> {csvFilePath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting {jsonFilePath}: {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}