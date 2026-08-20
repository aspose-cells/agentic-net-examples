// Title: C# Aspose.Cells utility to scan folders for encrypted Excel files and create a password status report
// Description: A .NET console app that recursively scans a directory, filters supported Excel extensions, uses Aspose.Cells FileFormatUtil.DetectFileFormat to check the IsEncrypted flag, and writes a CSV‑style text file reporting each file path with its encryption status or any detection error.
// Keywords: Aspose.Cells encrypted Excel detection | C# scan folder for password‑protected spreadsheets | FileFormatUtil IsEncrypted example | generate Excel encryption report .NET | list encrypted .xlsx files programmatically | Excel file encryption audit C# | bulk password status report Aspose
// Common Searches: How to detect password protection on Excel files with Aspose.Cells C# | C# program to find encrypted .xlsx files recursively | Create CSV report of encrypted Excel workbooks using Aspose | FileFormatUtil DetectFileFormat encryption check example | Scan directory for protected Excel files .NET
// Developer Intent: Determine which Excel workbooks in a given folder are encrypted and output a concise status report.
// Use Cases: Compliance audit of shared drives to list all password‑protected Excel files. | Pre‑processing step in a data‑migration pipeline to skip or flag encrypted spreadsheets. | Scheduled Windows task that generates daily logs of encryption status for newly added files.
// AI Prompts: Generate C# code that uses Aspose.Cells FileFormatUtil to detect encryption for every Excel file in a directory and writes the results to a CSV file. | Explain how to handle exceptions from DetectFileFormat and record error details in the output report. | Suggest improvements to include the encryption algorithm name or attempt decryption with a known password while keeping the original report format.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace ExcelEncryptionScanner
{
    // A .NET console app that recursively scans a directory, filters supported Excel extensions, uses Aspose.Cells FileFormatUtil.DetectFileFormat to check the IsEncrypted flag, and writes a CSV‑style text file reporting each file path with its encryption status or any detection error.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the directory to scan; use first argument or current directory if none provided
            string targetDirectory = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

            // Prepare a list to hold report lines
            List<string> reportLines = new List<string>();
            reportLines.Add("File Path,IsEncrypted");

            // Define the set of extensions that Aspose.Cells can handle
            HashSet<string> excelExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".csv"
            };

            // Scan all files recursively
            foreach (string filePath in Directory.GetFiles(targetDirectory, "*.*", SearchOption.AllDirectories))
            {
                if (!excelExtensions.Contains(Path.GetExtension(filePath)))
                    continue; // Skip non‑Excel files

                try
                {
                    // Detect file format and encryption status
                    FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                    bool isEncrypted = formatInfo.IsEncrypted;

                    // Add result to the report
                    reportLines.Add($"{filePath},{isEncrypted}");
                }
                catch (Exception ex)
                {
                    // If detection fails, record the error
                    reportLines.Add($"{filePath},Error:{ex.Message}");
                }
            }

            // Write the report to a text file in the target directory
            string reportPath = Path.Combine(targetDirectory, "PasswordStatusReport.txt");
            File.WriteAllLines(reportPath, reportLines);

            Console.WriteLine($"Password status report generated at: {reportPath}");
        }
    }
}
