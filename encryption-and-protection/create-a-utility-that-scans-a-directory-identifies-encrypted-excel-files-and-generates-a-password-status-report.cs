using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace ExcelEncryptionScanner
{
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the directory to scan; use current directory if none provided
            string directoryPath = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

            // Prepare the output report file path
            string reportPath = Path.Combine(directoryPath, "PasswordStatusReport.txt");

            // Open the report file for writing
            using (StreamWriter writer = new StreamWriter(reportPath))
            {
                // Write header line
                writer.WriteLine("File Path,Is Encrypted");

                // Get all files with typical Excel extensions in the directory (non‑recursive)
                var excelFiles = Directory.GetFiles(directoryPath)
                                          .Where(f => IsExcelFile(f));

                foreach (string filePath in excelFiles)
                {
                    try
                    {
                        // Detect file format and encryption status
                        FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
                        bool isEncrypted = info.IsEncrypted;

                        // Write result to the report
                        writer.WriteLine($"{filePath},{isEncrypted}");
                    }
                    catch (Exception ex)
                    {
                        // If detection fails, note the error in the report
                        writer.WriteLine($"{filePath},Error: {ex.Message}");
                    }
                }
            }

            Console.WriteLine($"Password status report generated at: {reportPath}");
        }

        // Helper method to identify Excel related file extensions
        private static bool IsExcelFile(string filePath)
        {
            string[] excelExtensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".xlt", ".xltx", ".xltm", ".ots" };
            string ext = Path.GetExtension(filePath);
            return excelExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase);
        }
    }
}