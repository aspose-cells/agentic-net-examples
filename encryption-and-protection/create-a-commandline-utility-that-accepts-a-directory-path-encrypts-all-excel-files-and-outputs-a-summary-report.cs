// Title: C# .NET CLI Tool to Batch‑Encrypt Excel Workbooks with Aspose.Cells and Generate a Summary Report
// Description: A command‑line utility written in C# that accepts a directory path, recursively scans for Excel files (.xlsx, .xls, .xlsm, .xlsb), skips files already password‑protected, applies a strong password via Aspose.Cells Workbook.Settings and SetEncryptionOptions, saves the workbooks in place, verifies encryption, and creates a detailed text report of successes, skips and errors.
// Keywords: Aspose.Cells | C# encrypt Excel files | batch Excel encryption | CLI Excel password protection | strong encryption .NET | Excel workbook encryption report | detect encrypted Excel C# | set encryption options Aspose | command line Excel security | GDPR Excel encryption | HIPAA Excel protection
// Common Searches: how to encrypt all Excel files in a folder using Aspose.Cells C# | batch encrypt Excel workbooks from the command line .NET | C# utility to generate encryption report for Excel files | skip already encrypted Excel files when processing C# | set strong password for multiple Excel workbooks Aspose
// Developer Intent: Encrypt every Excel workbook in a specified folder with a consistent password and produce a comprehensive log of the operation.
// Use Cases: Secure confidential spreadsheets before archiving by running the tool on the archive directory. | Automate regulatory compliance (e.g., GDPR, HIPAA) by encrypting all Excel files on a shared drive each night. | Create an audit trail that records which files were encrypted, which were already protected, and any processing errors.
// AI Prompts: Add a command‑line option to accept a custom password and allow the user to choose encryption strength (128‑bit or 256‑bit). | Write a PowerShell wrapper that executes the utility for multiple directories, merges the generated reports, and emails the consolidated summary. | Extend the program with a decryption mode that takes a password, decrypts matching files, and updates the report accordingly.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace ExcelEncryptionUtility
{
    // A command‑line utility written in C# that accepts a directory path, recursively scans for Excel files (.xlsx, .xls, .xlsm, .xlsb), skips files already password‑protected, applies a strong password via Aspose.Cells Workbook.Settings and SetEncryptionOptions, saves the workbooks in place, verifies encryption, and creates a detailed text report of successes, skips and errors.
    class Program
    {
        // Password used for encrypting workbooks
        private const string EncryptionPassword = "StrongPassword123";

        static void Main(string[] args)
        {
            // Validate command‑line arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExcelEncryptionUtility <directoryPath>");
                return;
            }

            string directoryPath = args[0];

            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"Error: Directory \"{directoryPath}\" does not exist.");
                return;
            }

            // Prepare a StringBuilder for the summary report
            StringBuilder reportBuilder = new StringBuilder();
            reportBuilder.AppendLine($"Encryption Report - {DateTime.Now}");
            reportBuilder.AppendLine($"Target Directory: {directoryPath}");
            reportBuilder.AppendLine();

            // Define supported Excel extensions
            string[] excelExtensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb" };

            // Enumerate all files with supported extensions
            foreach (string filePath in Directory.EnumerateFiles(directoryPath, "*.*", SearchOption.AllDirectories))
            {
                if (Array.IndexOf(excelExtensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                    continue; // Skip non‑Excel files

                try
                {
                    // Detect if the file is already encrypted
                    FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                    bool alreadyEncrypted = formatInfo.IsEncrypted;

                    if (alreadyEncrypted)
                    {
                        reportBuilder.AppendLine($"{Path.GetFileName(filePath)} - Already encrypted, skipped.");
                        continue;
                    }

                    // Load the workbook (no password needed because it's not encrypted)
                    Workbook workbook = new Workbook(filePath);

                    // Set the password to protect the workbook
                    workbook.Settings.Password = EncryptionPassword;

                    // Apply strong encryption options (optional but recommended)
                    workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                    // Save the workbook, overwriting the original file
                    workbook.Save(filePath);

                    // Verify encryption status after saving
                    FileFormatInfo postInfo = FileFormatUtil.DetectFileFormat(filePath);
                    bool encryptionSucceeded = postInfo.IsEncrypted;

                    reportBuilder.AppendLine($"{Path.GetFileName(filePath)} - Encryption {(encryptionSucceeded ? "succeeded" : "failed")}.");
                }
                catch (Exception ex)
                {
                    // Record any errors for this file
                    reportBuilder.AppendLine($"{Path.GetFileName(filePath)} - Error: {ex.Message}");
                }
            }

            // Output the report to console
            Console.WriteLine(reportBuilder.ToString());

            // Write the report to a text file in the target directory
            string reportPath = Path.Combine(directoryPath, "EncryptionReport.txt");
            File.WriteAllText(reportPath, reportBuilder.ToString());

            Console.WriteLine($"Report saved to: {reportPath}");
        }
    }
}
