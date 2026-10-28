// Title: C# batch conversion of JSON spreadsheets to CSV files using Aspose.Cells
// Common Searches: how to convert multiple JSON files to CSV with Aspose.Cells in a .NET console application | C# example loading a JSON workbook and saving it as CSV using Aspose.Cells | batch process all JSON spreadsheets in a folder to CSV using Aspose.Cells LoadOptions | automate JSON to CSV conversion for a directory of files in C# | Aspose.Cells script to transform JSON data into CSV format on Windows
// Tags: Aspose.Cells JSON to CSV batch processing | LoadOptions LoadFormat.Json usage | Workbook.Save SaveFormat.Csv example | C# directory file enumeration for conversion | robust error handling in Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// The C# console program enumerates every *.json file in a specified input folder, loads each file into an Aspose.Cells Workbook using LoadOptions with LoadFormat.Json, and saves it as a CSV file with the same base name in an output folder. It creates the output directory if needed and includes per‑file and global exception handling.
class Program
{
    static void Main()
    {
        // Directory containing JSON files
        string inputDirectory = @"C:\InputJson";
        // Directory where CSV files will be saved
        string outputDirectory = @"C:\OutputCsv";

        try
        {
            // Ensure the input directory exists
            if (!Directory.Exists(inputDirectory))
            {
                Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Retrieve all JSON files in the input directory
            string[] jsonFiles = Directory.GetFiles(inputDirectory, "*.json");

            foreach (string jsonFilePath in jsonFiles)
            {
                try
                {
                    // Verify the JSON file still exists
                    if (!File.Exists(jsonFilePath))
                    {
                        Console.WriteLine($"File not found: {jsonFilePath}");
                        continue;
                    }

                    // Load the JSON file into a workbook
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Json);
                    Workbook workbook = new Workbook(jsonFilePath, loadOptions);

                    // Build the corresponding CSV file path
                    string csvFileName = Path.GetFileNameWithoutExtension(jsonFilePath) + ".csv";
                    string csvFilePath = Path.Combine(outputDirectory, csvFileName);

                    // Save the workbook as a CSV file
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
