using System;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing Excel files
            string inputFolder = @"C:\InputExcelFiles";
            // Output folder where CSV files will be saved
            string outputFolder = @"C:\OutputCsvFiles";

            // Verify input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Define the Excel file extensions to process
            string[] excelExtensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb", ".ods" };

            try
            {
                // Enumerate all files with the defined extensions
                foreach (string filePath in Directory.EnumerateFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
                {
                    if (Array.IndexOf(excelExtensions, Path.GetExtension(filePath).ToLower()) < 0)
                        continue; // Skip non‑Excel files

                    // Verify the file still exists before processing
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook (default LoadOptions are sufficient for Excel formats)
                        Workbook workbook = new Workbook(filePath);

                        // Prepare CSV save options
                        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                        {
                            // Trim leading blank rows/columns (default is true, set explicitly for clarity)
                            TrimLeadingBlankRowAndColumn = true,
                            // Export all worksheets into a single CSV (set to true if required)
                            ExportAllSheets = true
                        };

                        // Build output CSV file name (same base name, .csv extension)
                        string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".csv";
                        string outputPath = Path.Combine(outputFolder, outputFileName);

                        // Save the workbook as CSV using the configured options
                        workbook.Save(outputPath, saveOptions);

                        Console.WriteLine($"Converted: {Path.GetFileName(filePath)} -> {outputFileName}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
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