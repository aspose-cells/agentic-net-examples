using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace JsonToCsvBatch
{
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing JSON files. Change as needed.
            string sourceDirectory = @"C:\JsonFiles";

            // Verify the directory exists.
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Directory not found: {sourceDirectory}");
                return;
            }

            // Get all JSON files in the directory.
            string[] jsonFiles = Directory.GetFiles(sourceDirectory, "*.json", SearchOption.TopDirectoryOnly);

            if (jsonFiles.Length == 0)
            {
                Console.WriteLine("No JSON files found to convert.");
                return;
            }

            foreach (string jsonPath in jsonFiles)
            {
                try
                {
                    // Determine the output CSV file path (same name, .csv extension).
                    string csvPath = Path.ChangeExtension(jsonPath, ".csv");

                    // Load the JSON file into a workbook.
                    JsonLoadOptions loadOptions = new JsonLoadOptions();
                    Workbook workbook = new Workbook(jsonPath, loadOptions);

                    // Save the workbook as CSV.
                    workbook.Save(csvPath, SaveFormat.Csv);

                    Console.WriteLine($"Converted: {Path.GetFileName(jsonPath)} -> {Path.GetFileName(csvPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{Path.GetFileName(jsonPath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}