// Title: Record each Excel workbook encryption toggle in a log file while using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that opens an .xlsx file with Aspose.Cells, checks Workbook.Settings.IsEncrypted, flips the encryption state by setting or clearing the password, saves the workbook, and appends a timestamped message to a specified log file. | Enhance existing Aspose.Cells code to capture both successful encryption toggles and exception details, ensuring the log file is created if missing and each entry includes the operation time and file name.
// Common Searches: asp.net toggle password protection on an Excel workbook using Aspose.Cells | c# log encryption status change of an .xlsx file with Aspose.Cells | how to programmatically enable or disable workbook encryption and write an audit log in .NET | aspose.cells save encrypted workbook and append operation log entry | detect Workbook.Settings.IsEncrypted and change password in C# console app
// Tags: toggle workbook encryption Aspose.Cells | audit log encryption changes C# | save encrypted Excel file Aspose.Cells | append timestamped entry to text log C# | modify Workbook.Settings.IsEncrypted programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptionToggle
{
    // // C# console program that loads Sample.xlsx with Aspose.Cells, checks Workbook.Settings.IsEncrypted, flips the encryption by setting or clearing the password, saves the result as Sample_Toggled.xlsx, and writes a timestamped line to EncryptionToggleLog.txt for each toggle or error.
    class Program
    {
        // Path to the workbook file
        private const string InputFilePath = @"C:\Data\Sample.xlsx";
        private const string OutputFilePath = @"C:\Data\Sample_Toggled.xlsx";

        // Log file path
        private const string LogFilePath = @"C:\Data\EncryptionToggleLog.txt";

        static void Main()
        {
            try
            {
                // Verify input file exists
                if (!File.Exists(InputFilePath))
                {
                    Console.Error.WriteLine($"Input file not found: {InputFilePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(InputFilePath);

                // Determine current encryption status (read‑only)
                bool isEncrypted = workbook.Settings.IsEncrypted;

                // Toggle encryption status
                if (isEncrypted)
                {
                    // Remove encryption by clearing the password
                    workbook.Settings.Password = string.Empty;
                    Log($"[{DateTime.Now}] Encryption disabled for workbook '{Path.GetFileName(InputFilePath)}'.");
                }
                else
                {
                    // Apply encryption with a password
                    workbook.Settings.Password = "StrongPassword123!";
                    Log($"[{DateTime.Now}] Encryption enabled for workbook '{Path.GetFileName(InputFilePath)}' with password.");
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(OutputFilePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (preserving the new encryption setting)
                workbook.Save(OutputFilePath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
                Log($"[{DateTime.Now}] Error: {ex.Message}");
            }
        }

        // Simple logger that appends messages to a text file
        private static void Log(string message)
        {
            try
            {
                File.AppendAllText(LogFilePath, message + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // In a real scenario, handle logging failures appropriately
                Console.Error.WriteLine($"Logging failed: {ex.Message}");
            }
        }
    }
}
