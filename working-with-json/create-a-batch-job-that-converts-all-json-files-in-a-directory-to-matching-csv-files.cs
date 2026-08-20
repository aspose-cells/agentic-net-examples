// Title: C# Batch Convert JSON Files to CSV Using Aspose.Cells
// Description: A console utility that scans a specified folder for *.json files, loads each file into an Aspose.Cells Workbook with JsonLoadOptions, and saves it as a CSV file with the same base name. Includes directory validation, per‑file error handling, and progress output.
// Keywords: Aspose.Cells JSON to CSV | C# batch JSON conversion | JsonLoadOptions | convert multiple JSON files to CSV | Aspose.Cells .NET CSV export | automated JSON to CSV conversion
// Common Searches: batch convert json to csv c# aspose.cells | convert all json files in a folder to csv using asp.net | aspocells jsonloadoptions example | c# program to export json as csv with aspose | automate json to csv conversion .net
// Developer Intent: Automatically transform every JSON file in a given directory into a CSV file with matching names using Aspose.Cells.
// Use Cases: Migrate exported JSON datasets to CSV for BI tools. | Process large collections of JSON logs into CSV for analytics pipelines. | Schedule nightly jobs that keep CSV mirrors of JSON configuration files.
// AI Prompts: Generate C# code that iterates through a folder, loads each JSON file into an Aspose.Cells Workbook via JsonLoadOptions, and saves it as CSV with robust error handling. | Rewrite the batch converter to use async I/O and Parallel.ForEach while preserving the Aspose.Cells workflow. | Create a PowerShell script that runs the compiled C# batch converter, captures success/failure messages, and writes a detailed log file.

using System;
using System.IO;
using Aspose.Cells;

namespace JsonToCsvBatch
{
    // A console utility that scans a specified folder for *.json files, loads each file into an Aspose.Cells Workbook with JsonLoadOptions, and saves it as a CSV file with the same base name. Includes directory validation, per‑file error handling, and progress output.
    class Program
    {
        static void Main()
        {
            // Directory containing JSON files
            string sourceDirectory = @"C:\Data\JsonFiles";

            // Verify that the source directory exists
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory not found: {sourceDirectory}");
                return;
            }

            try
            {
                // Get all JSON files in the directory
                string[] jsonFiles = Directory.GetFiles(sourceDirectory, "*.json");

                foreach (string jsonFilePath in jsonFiles)
                {
                    try
                    {
                        // Determine the output CSV file path (same name, .csv extension)
                        string csvFilePath = Path.ChangeExtension(jsonFilePath, ".csv");

                        // Load the JSON file into a workbook using JsonLoadOptions
                        JsonLoadOptions loadOptions = new JsonLoadOptions();
                        Workbook workbook = new Workbook(jsonFilePath, loadOptions);

                        // Save the workbook as CSV
                        workbook.Save(csvFilePath, SaveFormat.Csv);

                        Console.WriteLine($"Converted: {Path.GetFileName(jsonFilePath)} -> {Path.GetFileName(csvFilePath)}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting {Path.GetFileName(jsonFilePath)}: {ex.Message}");
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
