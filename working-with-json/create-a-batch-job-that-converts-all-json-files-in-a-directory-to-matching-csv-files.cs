// Title: Batch convert JSON files to CSV in C# using Aspose.Cells LoadOptions
// Common Searches: C# code to convert all JSON files in a folder to CSV with Aspose.Cells | How to process multiple JSON files and export them as CSV using Aspose.Cells .NET | Automate JSON to CSV conversion for a directory in C# with Aspose.Cells | Example of loading JSON workbooks and saving as CSV in bulk using Aspose.Cells | Aspose.Cells batch conversion of JSON files to CSV in a Windows environment
// Tags: Aspose.Cells LoadOptions LoadFormat.Json | Aspose.Cells SaveFormat.Csv export | C# iterate directory files Aspose.Cells | Workbook constructor with JSON source | per‑file exception handling Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// Scans a specified input directory for *.json files, loads each file into an Aspose.Cells Workbook using LoadOptions with LoadFormat.Json, and saves the workbook as a .csv file in an output directory. The program creates the output folder if it does not exist and logs per‑file errors without stopping the batch.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Input directory containing JSON files (first argument or default)
            string inputDir = args.Length > 0 ? args[0] : @"C:\InputJson";

            // Output directory for generated CSV files (second argument or default)
            string outputDir = args.Length > 1 ? args[1] : @"C:\OutputCsv";

            // Verify input directory exists
            if (!Directory.Exists(inputDir))
            {
                Console.WriteLine($"Input directory does not exist: {inputDir}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDir);

            // Retrieve all JSON files in the input directory (non‑recursive)
            string[] jsonFiles = Directory.GetFiles(inputDir, "*.json", SearchOption.TopDirectoryOnly);

            if (jsonFiles.Length == 0)
            {
                Console.WriteLine("No JSON files found to process.");
                return;
            }

            // Load options for JSON format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Json);

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

                    // Load the JSON file into a new workbook
                    Workbook workbook = new Workbook(jsonFilePath, loadOptions);

                    // Build the corresponding CSV file path
                    string csvFileName = Path.GetFileNameWithoutExtension(jsonFilePath) + ".csv";
                    string csvFilePath = Path.Combine(outputDir, csvFileName);

                    // Save the workbook as CSV
                    workbook.Save(csvFilePath, SaveFormat.Csv);

                    Console.WriteLine($"Converted '{jsonFilePath}' to '{csvFilePath}'.");
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
