// Title: C# program using Aspose.Cells to scan folders for password‑protected Excel files and export a CSV encryption status report
// Common Searches: how to list encrypted Excel files in a directory with Aspose.Cells C# | C# generate CSV report of password‑protected .xlsx files using Aspose.Cells | detect if an Excel workbook is password protected without providing a password in .NET | recursive scan for protected Excel workbooks and create status report in C# | Aspose.Cells check workbook encryption and export results to CSV
// Tags: Aspose.Cells detect encrypted workbook | C# scan directory for password protected Excel | generate CSV report of Excel encryption status | recursive file enumeration .xlsx Aspose.Cells | Workbook load without password exception handling | Excel file encryption detection using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ExcelPasswordStatusReport
{
    // The utility enumerates Excel files (.xls, .xlsx, .xlsm, .xlsb, .xls2003, .xls2007) in a specified or current folder, attempts to open each workbook with Aspose.Cells without a password, treats a CellsException containing "password" as an encrypted file, records the file path and encryption flag, and writes a PasswordStatusReport.csv containing FilePath and IsEncrypted columns.
    class Program
    {
        // Entry point of the utility
        static void Main(string[] args)
        {
            try
            {
                // Determine the target directory (first argument or current directory)
                string targetDirectory = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

                if (!Directory.Exists(targetDirectory))
                {
                    Console.WriteLine($"Directory does not exist: {targetDirectory}");
                    return;
                }

                // Supported Excel extensions
                string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".xls2003", ".xls2007" };

                // List to hold report entries
                var reportEntries = new List<ReportEntry>();

                // Scan the directory recursively for Excel files
                foreach (string filePath in Directory.EnumerateFiles(targetDirectory, "*.*", SearchOption.AllDirectories))
                {
                    if (!Array.Exists(extensions, ext => ext.Equals(Path.GetExtension(filePath), StringComparison.OrdinalIgnoreCase)))
                        continue;

                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    bool isEncrypted;
                    try
                    {
                        isEncrypted = IsFileEncrypted(filePath);
                    }
                    catch (Exception ex)
                    {
                        // Log unexpected errors and treat the file as not encrypted for reporting purposes
                        Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                        isEncrypted = false;
                    }

                    reportEntries.Add(new ReportEntry
                    {
                        FilePath = filePath,
                        IsEncrypted = isEncrypted
                    });
                }

                // Generate CSV report
                string reportPath = Path.Combine(targetDirectory, "PasswordStatusReport.csv");
                WriteCsvReport(reportPath, reportEntries);

                Console.WriteLine($"Report generated at: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }

        // Determines whether an Excel file is encrypted (password protected)
        private static bool IsFileEncrypted(string filePath)
        {
            try
            {
                // Attempt to load the workbook without providing a password
                // If the file is encrypted, Aspose.Cells will throw a CellsException
                var workbook = new Workbook(filePath);
                // Load succeeded, file is not encrypted
                return false;
            }
            catch (CellsException ex)
            {
                // Aspose.Cells uses the message containing "password" for encrypted files
                if (!string.IsNullOrEmpty(ex.Message) &&
                    ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                // Re‑throw for other CellsException types (e.g., corrupted file)
                throw;
            }
        }

        // Writes the collected report entries to a CSV file
        private static void WriteCsvReport(string reportPath, List<ReportEntry> entries)
        {
            using (var writer = new StreamWriter(reportPath, false))
            {
                // Header line
                writer.WriteLine("FilePath,IsEncrypted");

                // Data lines
                foreach (var entry in entries)
                {
                    // Escape double quotes in file path and wrap with quotes if needed
                    string escapedPath = entry.FilePath.Contains(",") || entry.FilePath.Contains("\"")
                        ? $"\"{entry.FilePath.Replace("\"", "\"\"")}\""
                        : entry.FilePath;

                    writer.WriteLine($"{escapedPath},{entry.IsEncrypted}");
                }
            }
        }

        // Simple DTO to hold report data for each file
        private class ReportEntry
        {
            public string FilePath { get; set; }
            public bool IsEncrypted { get; set; }
        }
    }
}
