using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace ExcelPasswordStatusReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Directory to scan – can be passed as first argument or default to current directory
            string directoryPath = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"Directory does not exist: {directoryPath}");
                return;
            }

            // Supported Excel file extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".csv", ".xml" };

            // List to hold report lines
            List<string> reportLines = new List<string>();
            reportLines.Add("File Path,Is Encrypted");

            // Scan files
            foreach (string filePath in Directory.EnumerateFiles(directoryPath, "*.*", SearchOption.AllDirectories))
            {
                if (Array.Exists(extensions, ext => ext.Equals(Path.GetExtension(filePath), StringComparison.OrdinalIgnoreCase)))
                {
                    bool isEncrypted = false;
                    try
                    {
                        // Detect file format and encryption status
                        FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
                        isEncrypted = info.IsEncrypted;
                    }
                    catch (Exception ex)
                    {
                        // If detection fails, treat as not encrypted and log the error
                        Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                    }

                    // Add result to report
                    reportLines.Add($"{filePath},{isEncrypted}");
                }
            }

            // Write report to a text file in the same directory
            string reportPath = Path.Combine(directoryPath, "PasswordStatusReport.txt");
            try
            {
                File.WriteAllLines(reportPath, reportLines);
                Console.WriteLine($"Password status report generated at: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write report: {ex.Message}");
            }
        }
    }
}