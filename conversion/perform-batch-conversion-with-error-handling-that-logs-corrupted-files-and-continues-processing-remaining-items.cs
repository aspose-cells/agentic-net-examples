// Title: Convert a folder of Excel workbooks to PDF, XLSX, CSV, or HTML with Aspose.Cells in C# and log corrupted files
// Common Searches: C# Aspose.Cells batch convert Excel files in a directory to PDF with error logging | How to skip corrupted Excel workbooks while converting a folder using Aspose.Cells | Aspose.Cells SaveFormat selection for batch converting Excel to CSV in C# | Log conversion failures to a file during bulk Excel to HTML conversion with Aspose.Cells | Ignore non‑Excel files when processing a folder with Aspose.Cells C#
// Tags: batch Excel to PDF conversion Aspose.Cells C# | folder workbook conversion SaveFormat Aspose.Cells | error handling and logging Aspose.Cells conversion | skip non‑Excel files Aspose.Cells batch processing | convert Excel to CSV Aspose.Cells C# | log corrupted workbook errors Aspose.Cells | automatic format detection Workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchConversion
{
    // The example scans a specified input folder for Excel files, loads each workbook with Aspose.Cells, saves it in the requested format (PDF, XLSX, CSV, HTML or default XLSX), logs any load or save exceptions to a conversion_errors.log file, and continues processing the remaining files while skipping non‑Excel items.
    public class BatchConverter
    {
        /// <param name="inputFolder">Folder containing source Excel files.</param>
        /// <param name="outputFolder">Folder where converted files will be saved.</param>
        /// <param name="outputFormat">Desired output format (e.g., "pdf", "xlsx", "csv").</param>
        public static void ConvertFolder(string inputFolder, string outputFolder, string outputFormat)
        {
            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all files with typical Excel extensions
            string[] files = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in files)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();

                // Skip non-Excel files
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".xlsb")
                    continue;

                // Ensure the file actually exists before attempting to load
                if (!File.Exists(filePath))
                {
                    string missingMsg = $"File not found: {filePath}";
                    Console.WriteLine(missingMsg);
                    LogError(missingMsg, outputFolder);
                    continue;
                }

                try
                {
                    // Load the workbook (Aspose.Cells automatically detects format)
                    Workbook workbook = new Workbook(filePath);

                    // Determine output file name and format
                    string outputFileName = Path.GetFileNameWithoutExtension(filePath) + "." + outputFormat;
                    string outputPath = Path.Combine(outputFolder, outputFileName);

                    // Save according to requested format
                    switch (outputFormat.ToLowerInvariant())
                    {
                        case "pdf":
                            workbook.Save(outputPath, SaveFormat.Pdf);
                            break;
                        case "xlsx":
                            workbook.Save(outputPath, SaveFormat.Xlsx);
                            break;
                        case "csv":
                            workbook.Save(outputPath, SaveFormat.Csv);
                            break;
                        case "html":
                            workbook.Save(outputPath, SaveFormat.Html);
                            break;
                        default:
                            // If format not recognized, default to Xlsx
                            workbook.Save(outputPath, SaveFormat.Xlsx);
                            break;
                    }

                    Console.WriteLine($"Successfully converted: {Path.GetFileName(filePath)} -> {outputFileName}");
                }
                catch (Exception ex)
                {
                    // Log corrupted or problematic files and continue
                    string logMessage = $"Error processing file '{filePath}': {ex.Message}";
                    Console.WriteLine(logMessage);
                    LogError(logMessage, outputFolder);
                }
            }
        }

        private static void LogError(string message, string outputFolder)
        {
            try
            {
                string logPath = Path.Combine(outputFolder, "conversion_errors.log");
                File.AppendAllText(logPath, $"{DateTime.Now:u} - {message}{Environment.NewLine}");
            }
            catch
            {
                // Swallow logging exceptions to keep batch processing alive
            }
        }
    }

    public class Program
    {
        // Entry point required for compilation
        public static void Main(string[] args)
        {
            // Expected arguments: inputFolder outputFolder outputFormat
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: AsposeCellsBatchConversion <inputFolder> <outputFolder> <outputFormat>");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];
            string outputFormat = args[2];

            // Validate input folder existence
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            try
            {
                BatchConverter.ConvertFolder(inputFolder, outputFolder, outputFormat);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
