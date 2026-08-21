// Title: Log Workbook Encryption Toggles with Aspose.Cells for .NET
// Description: C# example that uses Aspose.Cells to load an Excel file, checks the IsEncrypted flag, applies or removes a password, saves the workbook, and writes a timestamped entry for each change to a log file.
// Keywords: Aspose.Cells | C# | .NET | workbook encryption | toggle password protection | IsEncrypted property | audit log | Excel security | timestamped logging | batch processing
// Common Searches: Aspose.Cells log workbook encryption changes | C# toggle Excel password and record activity | how to track IsEncrypted flag with Aspose.Cells | write encryption audit log for Excel files | automated encryption/decryption logging .NET
// Developer Intent: Implement a reliable logger that records every encryption or decryption action performed on an Excel workbook using Aspose.Cells.
// Use Cases: Maintain an audit trail for compliance when passwords are added or removed from workbooks. | Detect missing protection in large‑scale Excel processing pipelines. | Provide change history for files in automated ETL or document‑management workflows.
// AI Prompts: Create C# code with Aspose.Cells that toggles workbook encryption and appends a timestamped log entry for each operation. | Refactor the ToggleEncryption method to use async file I/O and ensure thread‑safe logging across multiple files. | Explain how to extend the logger to capture previous and new encryption states and handle parallel processing of workbooks.

using System;
using System.IO;
using Aspose.Cells;

namespace EncryptionToggleLogger
{
    // Simple logger that writes entries to a text file.
    // C# example that uses Aspose.Cells to load an Excel file, checks the IsEncrypted flag, applies or removes a password, saves the workbook, and writes a timestamped entry for each change to a log file.
    public static class Logger
    {
        private static readonly string LogFilePath = "encryption_log.txt";

        public static void Log(string message)
        {
            string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            File.AppendAllLines(LogFilePath, new[] { entry });
        }
    }

    public class WorkbookEncryptionProcessor
    {
        // Toggles encryption on the given workbook.
        // If the workbook is currently encrypted, it removes the password.
        // If it is not encrypted, it applies the supplied password.
        // After each change, a log entry is created.
        public void ToggleEncryption(string workbookPath, string password)
        {
            if (!File.Exists(workbookPath))
            {
                Logger.Log($"File not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook with the supplied password (required if the file is encrypted)
                var loadOptions = new LoadOptions { Password = password };
                Workbook wb = new Workbook(workbookPath, loadOptions);

                bool wasEncrypted = wb.Settings.IsEncrypted;
                Logger.Log($"Loaded workbook '{Path.GetFileName(workbookPath)}'. Initial IsEncrypted = {wasEncrypted}");

                if (wasEncrypted)
                {
                    // Remove encryption by clearing the password
                    wb.Settings.Password = string.Empty;
                    Logger.Log("Password cleared to remove encryption.");
                }
                else
                {
                    // Apply encryption by setting a password
                    wb.Settings.Password = password;
                    Logger.Log($"Password set to '{password}' to enable encryption.");
                }

                // Save the workbook (overwrites the original file)
                wb.Save(workbookPath);
                bool isNowEncrypted = wb.Settings.IsEncrypted;
                Logger.Log($"Workbook saved. New IsEncrypted = {isNowEncrypted}");
            }
            catch (Exception ex)
            {
                Logger.Log($"Error processing workbook '{Path.GetFileName(workbookPath)}': {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            // Path to the workbook to process
            string filePath = "SampleWorkbook.xlsx";

            // Ensure a workbook exists (create rule)
            if (!File.Exists(filePath))
            {
                try
                {
                    Workbook newWb = new Workbook();
                    newWb.Worksheets[0].Cells["A1"].PutValue("Demo data");
                    newWb.Save(filePath);
                    Logger.Log($"Created new workbook at '{filePath}'.");
                }
                catch (Exception ex)
                {
                    Logger.Log($"Failed to create workbook: {ex.Message}");
                    return;
                }
            }

            var processor = new WorkbookEncryptionProcessor();

            // First toggle: encrypt the workbook
            processor.ToggleEncryption(filePath, "MySecretPwd");

            // Second toggle: decrypt the workbook
            processor.ToggleEncryption(filePath, "MySecretPwd");
        }
    }
}
