// Title: Batch convert Excel workbooks to PDF or other formats in C# using Aspose.Cells with error logging
// Common Searches: how to batch convert all .xlsx files in a directory to PDF using Aspose.Cells in C# while ignoring corrupted workbooks | C# Aspose.Cells example for converting multiple Excel files to different formats with try‑catch logging | skip non‑Excel files during Aspose.Cells batch processing and write errors to a log file | map Aspose.Cells SaveFormat enum to appropriate file extensions in a C# batch converter | continue processing after a workbook load exception when using Aspose.Cells in C#
// Tags: Aspose.Cells batch workbook conversion C# | Excel to PDF conversion with error logging Aspose.Cells | SaveFormat enum file extension mapping Aspose.Cells | skip unsupported Excel extensions during batch conversion | log failed workbook conversions Aspose.Cells C#

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsBatchConversion
{
    // The BatchConverter class recursively scans a source folder for supported Excel files, loads each workbook with Aspose.Cells, saves it in the chosen SaveFormat (e.g., PDF), and logs any load or save exceptions to a file while continuing to process the remaining documents.
    public class BatchConverter
    {
        // Path to the folder containing source Excel files
        private readonly string _sourceFolder;
        // Path to the folder where converted files will be saved
        private readonly string _outputFolder;
        // Desired output format (e.g., Pdf, Html, Csv, etc.)
        private readonly SaveFormat _targetFormat;
        // Log file for corrupted or failed files
        private readonly string _logFilePath;

        public BatchConverter(string sourceFolder, string outputFolder, SaveFormat targetFormat)
        {
            _sourceFolder = sourceFolder;
            _outputFolder = outputFolder;
            _targetFormat = targetFormat;
            _logFilePath = Path.Combine(_outputFolder, "conversion_errors.log");

            // Ensure output directory exists
            Directory.CreateDirectory(_outputFolder);
            // Clear previous log
            if (File.Exists(_logFilePath))
                File.Delete(_logFilePath);
        }

        public void ConvertAll()
        {
            // Verify source folder exists
            if (!Directory.Exists(_sourceFolder))
            {
                LogError(_sourceFolder, new DirectoryNotFoundException("Source folder does not exist."));
                return;
            }

            // Get all Excel files in the source folder (including subfolders)
            string[] excelFiles = Directory.GetFiles(_sourceFolder, "*.*", SearchOption.AllDirectories);
            var supportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv", ".ods"
            };

            foreach (string filePath in excelFiles)
            {
                if (!supportedExtensions.Contains(Path.GetExtension(filePath)))
                    continue; // Skip non‑Excel files

                // Ensure the file still exists before processing
                if (!File.Exists(filePath))
                {
                    LogError(filePath, new FileNotFoundException("File not found.", filePath));
                    continue;
                }

                try
                {
                    // Load the workbook (Aspose.Cells automatically detects format)
                    Workbook workbook = new Workbook(filePath);

                    // Determine output file name and extension based on target format
                    string outputFileName = Path.GetFileNameWithoutExtension(filePath) + GetExtensionForFormat(_targetFormat);
                    string outputPath = Path.Combine(_outputFolder, outputFileName);

                    // Save the workbook in the desired format
                    workbook.Save(outputPath, _targetFormat);
                }
                catch (Exception ex)
                {
                    // Log the error and continue with the next file
                    LogError(filePath, ex);
                }
            }
        }

        private void LogError(string filePath, Exception ex)
        {
            string message = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Failed to convert \"{filePath}\": {ex.Message}";
            Console.WriteLine(message);
            try
            {
                File.AppendAllLines(_logFilePath, new[] { message });
            }
            catch
            {
                // Swallow logging errors to keep batch processing alive
            }
        }

        private string GetExtensionForFormat(SaveFormat format)
        {
            // Map SaveFormat to typical file extensions
            switch (format)
            {
                case SaveFormat.Pdf:
                    return ".pdf";
                case SaveFormat.Html:
                    return ".html";
                case SaveFormat.Csv:
                    return ".csv";
                case SaveFormat.Xlsx:
                    return ".xlsx";
                case SaveFormat.Xlsb:
                    return ".xlsb";
                case SaveFormat.Xlsm:
                    return ".xlsm";
                case SaveFormat.Ods:
                    return ".ods";
                // SaveFormat.Xls is not available in this version; treat it as .xls via default case
                default:
                    return ".out";
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main(string[] args)
        {
            // Adjust these paths as needed
            string sourceFolder = @"C:\InputExcelFiles";
            string outputFolder = @"C:\ConvertedFiles";

            // Choose the target format (e.g., PDF)
            SaveFormat targetFormat = SaveFormat.Pdf;

            var converter = new BatchConverter(sourceFolder, outputFolder, targetFormat);
            converter.ConvertAll();

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
