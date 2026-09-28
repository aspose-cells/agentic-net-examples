// Title: C# batch conversion of Excel workbooks to PDF (or other formats) using Aspose.Cells with error logging and resilient processing
// AI Prompts: Create a C# method that scans a directory for .xls, .xlsx, .xlsm, and .xlsb files, loads each workbook with Aspose.Cells, saves it to a specified SaveFormat, and appends any exception details to a log file while keeping the loop alive. | Enhance the converter to produce a post‑run summary that reports the count of successful conversions, failed conversions, and the location of the error log. | Implement a retry strategy that re‑attempts a failed workbook conversion up to two additional times for transient I/O errors, preserving the existing logging behavior.
// Common Searches: how to use Aspose.Cells in C# to convert multiple Excel files to PDF and log errors | c# batch convert xls and xlsx files to pdf with Aspose.Cells while continuing on failure | asp.net core process a folder of Excel workbooks with Aspose.Cells and generate an error log | skip non‑excel files during bulk conversion using Aspose.Cells C# example | map Aspose.Cells SaveFormat to file extension in a batch conversion script
// Tags: Aspose.Cells bulk workbook conversion | C# Excel to PDF batch processing | error logging for Aspose.Cells conversion | SaveFormat to file extension mapping C# | skip unsupported file types Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsBatchConversion
{
    // Processes all supported Excel files in a given folder, converts each to a chosen format with Aspose.Cells, logs any conversion errors, skips non‑Excel files, and continues processing remaining workbooks.
    public class BatchConverter
    {
        private readonly string _inputFolder;
        private readonly string _outputFolder;
        private readonly SaveFormat _targetFormat;
        private readonly string _logFilePath;

        /// <summary>
        /// Initializes a new instance of the BatchConverter class.
        /// </summary>
        /// <param name="inputFolder">Folder containing source Excel files.</param>
        /// <param name="outputFolder">Folder where converted files will be saved.</param>
        /// <param name="targetFormat">Desired output format (e.g., SaveFormat.Pdf).</param>
        public BatchConverter(string inputFolder, string outputFolder, SaveFormat targetFormat)
        {
            _inputFolder = inputFolder;
            _outputFolder = outputFolder;
            _targetFormat = targetFormat;
            _logFilePath = Path.Combine(_outputFolder, "conversion_errors.log");

            // Ensure output directory exists
            Directory.CreateDirectory(_outputFolder);
        }

        /// <summary>
        /// Executes the batch conversion.
        /// </summary>
        public void ConvertAll()
        {
            // Collect all Excel files (xls, xlsx, xlsm) in the input folder
            string[] excelFiles = Directory.GetFiles(_inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            List<string> supportedExtensions = new List<string> { ".xls", ".xlsx", ".xlsm", ".xlsb" };

            foreach (string filePath in excelFiles)
            {
                if (!supportedExtensions.Contains(Path.GetExtension(filePath), StringComparer.OrdinalIgnoreCase))
                {
                    // Skip non‑Excel files
                    continue;
                }

                try
                {
                    // Load the workbook (Aspose.Cells handles many Excel formats)
                    Workbook workbook = new Workbook(filePath);

                    // Determine output file name with appropriate extension
                    string outputFileName = Path.GetFileNameWithoutExtension(filePath) + GetExtensionForFormat(_targetFormat);
                    string outputPath = Path.Combine(_outputFolder, outputFileName);

                    // Save the workbook in the target format
                    workbook.Save(outputPath, _targetFormat);
                }
                catch (Exception ex)
                {
                    // Log the error and continue with the next file
                    LogError(filePath, ex);
                }
            }
        }

        /// <summary>
        /// Returns the file extension associated with a given SaveFormat.
        /// </summary>
        private string GetExtensionForFormat(SaveFormat format)
        {
            switch (format)
            {
                case SaveFormat.Pdf:
                    return ".pdf";
                case SaveFormat.Html:
                    return ".html";
                case SaveFormat.Csv:
                    return ".csv";
                case SaveFormat.Xps:
                    return ".xps";
                // Add more mappings as needed
                default:
                    return ".out";
            }
        }

        /// <summary>
        /// Appends an error entry to the log file.
        /// </summary>
        private void LogError(string filePath, Exception ex)
        {
            string message = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Failed to convert '{filePath}'. Error: {ex.Message}";
            try
            {
                File.AppendAllText(_logFilePath, message + Environment.NewLine);
            }
            catch
            {
                // If logging fails, fall back to console output
                Console.Error.WriteLine(message);
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
            string destinationFolder = @"C:\ConvertedFiles";

            // Create a converter that transforms Excel files to PDF
            var converter = new BatchConverter(sourceFolder, destinationFolder, SaveFormat.Pdf);
            converter.ConvertAll();

            Console.WriteLine("Batch conversion completed. Check the log file for any errors.");
        }
    }
}
