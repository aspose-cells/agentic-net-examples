// Title: Batch load Excel, CSV, and JSON files with identical LoadOptions using Aspose.Cells for .NET
// Description: Iterate through a folder, filter supported .xlsx, .xls, .csv, and .json files, create matching LoadOptions (AutoFilter enabled, formula parsing disabled), load each workbook, log worksheet count, and save the result as a new XLSX file in an output directory.
// Keywords: Aspose.Cells batch load | C# load multiple Excel files | LoadOptions AutoFilter | disable formula parsing Aspose.Cells | convert CSV to XLSX Aspose | process JSON Excel data | directory workbook processing | Aspose.Cells SaveFormat.Xlsx
// Common Searches: load all Excel files in a folder with Aspose.Cells | batch convert CSV and JSON to XLSX using Aspose.Cells | apply same LoadOptions to multiple workbooks C# | Aspose.Cells auto‑filter and skip formula parsing | process directory of mixed Excel formats Aspose
// Developer Intent: Load each workbook in a directory with uniform LoadOptions and save the processed files.
// Use Cases: Speed up bulk loading by disabling formula parsing while keeping auto‑filter active. | Standardize a mixed collection of .xlsx, .xls, .csv, and .json files into a single XLSX format. | Generate a quick worksheet‑count report for every file before further analysis.
// AI Prompts: Write a C# method that scans a directory, loads every .xlsx, .xls, .csv, and .json file with LoadOptions.AutoFilter = true and ParsingFormulaOnOpen = false, then saves each as a new XLSX file. | Provide Aspose.Cells code to batch convert CSV and JSON workbooks to XLSX while applying identical LoadOptions across all formats. | Create an example that logs the number of worksheets for each workbook during batch processing with shared LoadOptions.

using System;
using System.IO;
using Aspose.Cells;

namespace MyExcelBatch
{
    // Iterate through a folder, filter supported .xlsx, .xls, .csv, and .json files, create matching LoadOptions (AutoFilter enabled, formula parsing disabled), load each workbook, log worksheet count, and save the result as a new XLSX file in an output directory.
    public class BatchExcelProcessor
    {
        // Processes all supported Excel files in the input directory,
        // applies identical LoadOptions to each workbook, and saves the result.
        public static void ProcessDirectory(string inputDirectory, string outputDirectory)
        {
            try
            {
                if (!Directory.Exists(inputDirectory))
                {
                    Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                    return;
                }

                // Ensure the output directory exists.
                Directory.CreateDirectory(outputDirectory);

                // Get all files in the input directory.
                string[] allFiles = Directory.GetFiles(inputDirectory, "*.*", SearchOption.TopDirectoryOnly);

                foreach (string filePath in allFiles)
                {
                    string extension = Path.GetExtension(filePath).ToLowerInvariant();

                    // Process only supported Excel-related formats.
                    if (extension == ".xlsx" ||
                        extension == ".xls" ||
                        extension == ".csv" ||
                        extension == ".json")
                    {
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found: {filePath}");
                            continue;
                        }

                        // Create identical LoadOptions for each file, using the constructor that matches the format.
                        LoadOptions loadOptions;
                        switch (extension)
                        {
                            case ".xlsx":
                                loadOptions = new LoadOptions(LoadFormat.Xlsx);
                                break;
                            case ".xls":
                                loadOptions = new LoadOptions(LoadFormat.Excel97To2003);
                                break;
                            case ".csv":
                                loadOptions = new LoadOptions(LoadFormat.Csv);
                                break;
                            case ".json":
                                loadOptions = new LoadOptions(LoadFormat.Json);
                                break;
                            default:
                                loadOptions = new LoadOptions(); // Auto detection as fallback.
                                break;
                        }

                        // Set common options that should be applied to every workbook.
                        loadOptions.AutoFilter = true;                 // Enable auto‑filtering.
                        loadOptions.ParsingFormulaOnOpen = false;     // Skip formula parsing for speed.

                        // Load the workbook with the specified LoadOptions.
                        Workbook workbook = new Workbook(filePath, loadOptions);

                        // Example operation: output the number of worksheets.
                        Console.WriteLine($"Loaded '{Path.GetFileName(filePath)}' – Worksheets: {workbook.Worksheets.Count}");

                        // Save the processed workbook to the output directory in XLSX format.
                        string outputFileName = Path.GetFileNameWithoutExtension(filePath) + "_processed.xlsx";
                        string outputPath = Path.Combine(outputDirectory, outputFileName);
                        workbook.Save(outputPath, SaveFormat.Xlsx);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing directory: {ex.Message}");
            }
        }
    }

    public class Program
    {
        // Entry point required for console application.
        public static void Main(string[] args)
        {
            try
            {
                string inputDir = args.Length > 0 ? args[0] : @"C:\InputExcelFiles";
                string outputDir = args.Length > 1 ? args[1] : @"C:\ProcessedExcelFiles";

                BatchExcelProcessor.ProcessDirectory(inputDir, outputDir);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
