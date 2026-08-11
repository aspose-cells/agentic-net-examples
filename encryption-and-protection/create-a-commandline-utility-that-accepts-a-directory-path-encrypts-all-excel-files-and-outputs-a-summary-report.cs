// Title: C# CLI tool to batch‑encrypt Excel workbooks with Aspose.Cells and create a summary report
// Description: A console application that receives a folder path, scans for .xls, .xlsx, .xlsb and .xlsm files, skips those already password‑protected, applies a default password using Aspose.Cells, saves the workbooks in place, verifies encryption, and writes a detailed report to the console and to EncryptionReport.txt in the target directory.
// Keywords: Aspose.Cells encrypt Excel C# | batch Excel password protection | C# command line Excel encryption | detect encrypted workbook Aspose | Excel encryption summary report
// Common Searches: C# program to encrypt all Excel files in a folder | Aspose.Cells command line password protection example | how to generate encryption report for Excel workbooks | skip already encrypted Excel files C#
// Developer Intent: Secure every Excel file in a specified directory with a default password and produce a concise audit log of the operation.
// Use Cases: Mass‑protect confidential spreadsheets before uploading to a shared drive. | Automate compliance checks by ensuring all departmental Excel files are password‑locked. | Maintain an audit trail that records encrypted, skipped, and failed files for governance reporting.
// AI Prompts: Generate a C# method that encrypts a workbook with a given password using Aspose.Cells and returns true on success. | Modify the utility to walk subdirectories recursively and accept a custom password argument from the command line. | Explain why FileFormatUtil.DetectFileFormat is used to verify encryption status after saving a workbook.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace ExcelEncryptor
{
    // A console application that receives a folder path, scans for .xls, .xlsx, .xlsb and .xlsm files, skips those already password‑protected, applies a default password using Aspose.Cells, saves the workbooks in place, verifies encryption, and writes a detailed report to the console and to EncryptionReport.txt in the target directory.
    class Program
    {
        // Default password used for encryption
        private const string DefaultPassword = "Password123";

        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExcelEncryptor <directoryPath>");
                return;
            }

            string directoryPath = args[0];

            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"Error: Directory '{directoryPath}' does not exist.");
                return;
            }

            // Supported Excel extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsb", ".xlsm" };

            // Collect summary information
            List<string> reportLines = new List<string>();
            reportLines.Add($"Encryption Report - {DateTime.Now}");
            reportLines.Add($"Target Directory: {directoryPath}");
            reportLines.Add("");

            // Process each Excel file in the directory (non-recursive)
            foreach (string filePath in Directory.GetFiles(directoryPath))
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue; // Skip non-Excel files

                try
                {
                    // Detect if the file is already encrypted
                    FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                    bool alreadyEncrypted = formatInfo.IsEncrypted;

                    if (alreadyEncrypted)
                    {
                        reportLines.Add($"{Path.GetFileName(filePath)} - Already encrypted, skipped.");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Set password protection
                    workbook.Settings.Password = DefaultPassword;

                    // Optionally set stronger encryption options (ignored for .xlsx/.xlsm but harmless)
                    workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                    // Save back to the same file (overwrites original)
                    workbook.Save(filePath);

                    // Verify encryption status after saving
                    FileFormatInfo postInfo = FileFormatUtil.DetectFileFormat(filePath);
                    bool isNowEncrypted = postInfo.IsEncrypted;

                    reportLines.Add($"{Path.GetFileName(filePath)} - Encryption {(isNowEncrypted ? "succeeded" : "failed")}.");
                }
                catch (Exception ex)
                {
                    reportLines.Add($"{Path.GetFileName(filePath)} - Error: {ex.Message}");
                }
            }

            // Output the summary report to console
            Console.WriteLine();
            foreach (string line in reportLines)
            {
                Console.WriteLine(line);
            }

            // Optionally write the report to a text file in the target directory
            string reportPath = Path.Combine(directoryPath, "EncryptionReport.txt");
            File.WriteAllLines(reportPath, reportLines);
            Console.WriteLine();
            Console.WriteLine($"Report saved to: {reportPath}");
        }
    }
}
