// Title: Batch Convert Excel Files to CSV with Progress Reporting using Aspose.Cells in C#
// Description: Scans a folder for .xls, .xlsx, .xlsm, .xlsb, .ods files, creates an output directory if needed, and converts each workbook to CSV with Aspose.Cells.Utility.ConversionUtility while printing a progress line and handling errors.
// Keywords: Aspose.Cells | C# batch Excel to CSV | ConversionUtility | folder conversion | progress reporting | .NET CSV export | automated spreadsheet conversion | Windows console utility
// Common Searches: C# convert all Excel files in a folder to CSV Aspose.Cells | batch Excel to CSV conversion with progress messages | Aspose.Cells ConversionUtility example for multiple files | how to export xlsx files to csv programmatically .NET | command line tool to convert spreadsheets to csv
// Developer Intent: Convert every supported Excel workbook in a directory to CSV while showing per‑file progress.
// Use Cases: Nightly automation that turns report workbooks into CSV for downstream analytics. | Bulk migration of legacy .xls/.xlsx archives to CSV before loading into a data warehouse. | Command‑line utility for end users to generate CSV equivalents of a folder’s spreadsheets with status feedback.
// AI Prompts: Write a C# method that receives a list of Excel paths and uses Aspose.Cells ConversionUtility to produce CSV files, logging success or failure for each entry. | Create unit tests for ExcelToCsvBatchConverter.Run covering missing input folder, empty source directory, and simulated conversion exceptions. | Generate a PowerShell wrapper that runs the compiled batch converter executable, passing input and output directories as parameters.

using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    // Scans a folder for .xls, .xlsx, .xlsm, .xlsb, .ods files, creates an output directory if needed, and converts each workbook to CSV with Aspose.Cells.Utility.ConversionUtility while printing a progress line and handling errors.
    public class ExcelToCsvBatchConverter
    {
        /// <param name="inputFolder">Folder containing source Excel files.</param>
        /// <param name="outputFolder">Folder where CSV files will be saved.</param>
        public static void Run(string inputFolder, string outputFolder)
        {
            // Ensure the input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Create the output folder if it does not exist
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all Excel files (xls, xlsx, xlsm) in the input folder
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            var supportedExtensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".csv" };
            var filesToConvert = Array.FindAll(excelFiles, f => Array.Exists(supportedExtensions,
                ext => string.Equals(Path.GetExtension(f), ext, StringComparison.OrdinalIgnoreCase)));

            int total = filesToConvert.Length;
            if (total == 0)
            {
                Console.WriteLine("No Excel files found to convert.");
                return;
            }

            Console.WriteLine($"Found {total} file(s) to convert.");

            for (int i = 0; i < total; i++)
            {
                string sourcePath = filesToConvert[i];
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".csv");

                try
                {
                    // Use the provided ConversionUtility method to perform conversion
                    ConversionUtility.Convert(sourcePath, destPath);
                    Console.WriteLine($"[{i + 1}/{total}] Converted: {Path.GetFileName(sourcePath)} -> {Path.GetFileName(destPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[{i + 1}/{total}] Error converting {Path.GetFileName(sourcePath)}: {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string inputFolder = @"C:\InputExcelFiles";
            string outputFolder = @"C:\OutputCsvFiles";

            ExcelToCsvBatchConverter.Run(inputFolder, outputFolder);
        }
    }
}
