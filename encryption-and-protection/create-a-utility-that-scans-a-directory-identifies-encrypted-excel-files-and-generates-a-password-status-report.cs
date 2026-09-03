// Title: C# console utility to scan directories, detect password‑protected Excel files with Aspose.Cells, and generate a CSV password status report
// AI Prompts: Write a C# console program that recursively enumerates .xls, .xlsx, .xlsm, and .xlsb files, attempts to open each workbook with Aspose.Cells LoadOptions without supplying a password, and writes "Encrypted" or "Not Encrypted" for each file to a PasswordStatusReport.csv file. | Extend the utility to accept an optional command‑line argument specifying the output CSV location and add a column that records the specific protection type (e.g., password, read‑only, structure) when Aspose.Cells can determine it. | Add robust error handling that logs files causing non‑password exceptions to a separate error log and excludes those entries from the final CSV report.
// Common Searches: how to list encrypted Excel workbooks in a folder using Aspose.Cells C# | c# program to generate CSV of Excel files password protection status | detect password protected .xlsx files without providing a password using Aspose.Cells | recursive directory scan for Excel files and check encryption with Aspose.Cells | Aspose.Cells load workbook without password to determine if it is encrypted in C#
// Tags: identify encrypted Excel workbooks Aspose.Cells | recursive folder scan for Excel password protection C# | export workbook encryption status to CSV C# | load Excel file without password to test encryption Aspose.Cells | generate password status report for .xls .xlsx files

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelPasswordStatusReport
{
    // The sample program walks through a given (or current) directory and all subfolders, finds Excel files (.xls, .xlsx, .xlsm, .xlsb), attempts to open each workbook with Aspose.Cells without a password, determines whether the file is encrypted based on the thrown exception, and writes the file name and encryption status to a PasswordStatusReport.csv file.
    class Program
    {
        // Entry point of the utility
        static void Main(string[] args)
        {
            try
            {
                // Directory to scan – can be passed as first argument, otherwise use current directory
                string targetDirectory = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

                // Prepare the report file path
                string reportPath = Path.Combine(targetDirectory, "PasswordStatusReport.csv");

                // Write header line to the CSV report
                using (var writer = new StreamWriter(reportPath, false))
                {
                    writer.WriteLine("FileName,Status");
                }

                // Get all Excel files (xls, xlsx, xlsm, etc.) in the directory and subdirectories
                string[] excelFiles = Directory.GetFiles(targetDirectory, "*.*", SearchOption.AllDirectories);
                foreach (string filePath in excelFiles)
                {
                    string extension = Path.GetExtension(filePath).ToLowerInvariant();
                    if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".xlsb")
                    {
                        continue; // Skip non‑Excel files
                    }

                    // Ensure the file exists before attempting to load it
                    if (!File.Exists(filePath))
                    {
                        continue;
                    }

                    bool isEncrypted;
                    try
                    {
                        // Determine if the file is encrypted (password protected)
                        isEncrypted = IsExcelFileEncrypted(filePath);
                    }
                    catch (Exception ex)
                    {
                        // Log the error and treat the file as unknown (skip it)
                        Console.Error.WriteLine($"Error processing '{filePath}': {ex.Message}");
                        continue;
                    }

                    // Append result to the CSV report
                    using (var writer = new StreamWriter(reportPath, true))
                    {
                        writer.WriteLine($"{Path.GetFileName(filePath)},{(isEncrypted ? "Encrypted" : "Not Encrypted")}");
                    }
                }

                Console.WriteLine($"Password status report generated at: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Fatal error: {ex.Message}");
            }
        }

        // Checks whether an Excel file is password protected.
        // Returns true if the file is encrypted, false otherwise.
        private static bool IsExcelFileEncrypted(string filePath)
        {
            try
            {
                // Attempt to load the workbook without providing a password.
                // If the file is encrypted, Aspose.Cells will throw an exception.
                var loadOptions = new LoadOptions(LoadFormat.Auto);
                var workbook = new Workbook(filePath, loadOptions);

                // If loading succeeds, the file is not encrypted.
                return false;
            }
            catch (Exception ex)
            {
                // Aspose.Cells throws an exception whose message contains the word "password"
                // when the file is password protected.
                if (!string.IsNullOrEmpty(ex.Message) &&
                    ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                // For any other exception, rethrow as it indicates a different problem.
                throw;
            }
        }
    }
}
