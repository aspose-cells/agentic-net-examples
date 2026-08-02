// Title: Timestamped logging of workbook encryption actions with Aspose.Cells for .NET
// Description: Demonstrates a static `EncryptionLogger` that appends a line with the current date‑time, workbook file name, and a descriptive action to a text file. The sample creates a workbook, assigns a password, sets encryption options, saves the file, and reloads it for verification, logging each step to provide an audit trail of encryption operations.
// Keywords: Aspose.Cells | .NET | C# | encryption logging | workbook protection audit | timestamp log | Excel encryption | password protected workbook | log file output | LoadOptions password | SaveFormat Xlsx | EncryptionType StrongCryptographicProvider | security event tracking
// Common Searches: Aspose.Cells log encryption actions | C# timestamp log for Excel workbook protection | how to audit password protection with Aspose.Cells | record workbook encryption steps to a file .NET | Aspose.Cells encryption logger example
// Developer Intent: Create a reusable logger that records the timestamp, file name, and each encryption‑related action performed on an Aspose.Cells workbook.
// Use Cases: Track when a password is applied to a workbook for compliance reporting. | Record the selection of encryption type and key size before saving the file. | Log successful save and verification loads to detect tampering or failures. | Generate an audit trail for batch processing of multiple workbooks.
// AI Prompts: Generate a thread‑safe Aspose.Cells encryption logger class with configurable log file path. | Provide code to extend the logger to capture operation duration and handle exceptions during encryption. | Show how to integrate the logger into a loop that processes dozens of workbooks, ensuring each encryption step is recorded.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionLogging
{
    // Simple logger that writes timestamp, file name and action to a text file
    // Demonstrates a static `EncryptionLogger` that appends a line with the current date‑time, workbook file name, and a descriptive action to a text file. The sample creates a workbook, assigns a password, sets encryption options, saves the file, and reloads it for verification, logging each step to provide an audit trail of encryption operations.
    public static class EncryptionLogger
    {
        private static readonly string LogFilePath = "encryption_log.txt";

        public static void Log(string fileName, string action)
        {
            string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\t{fileName}\t{action}";
            File.AppendAllLines(LogFilePath, new[] { entry });
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Define the output file name
            string outputFile = "EncryptedWorkbook.xlsx";

            // Set a password to protect the workbook
            string password = "StrongPassword123";
            workbook.Settings.Password = password;
            EncryptionLogger.Log(outputFile, $"Password set to protect workbook (Password: {password})");

            // Set encryption options (optional for Excel 2003 compatibility)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);
            EncryptionLogger.Log(outputFile, "Encryption options set: StrongCryptographicProvider, 128-bit key");

            // Save the encrypted workbook
            workbook.Save(outputFile, SaveFormat.Xlsx);
            EncryptionLogger.Log(outputFile, "Workbook saved with encryption");

            // Verify by loading the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions { Password = password };
            Workbook loadedWorkbook = new Workbook(outputFile, loadOptions);
            EncryptionLogger.Log(outputFile, "Encrypted workbook loaded successfully for verification");

            // Clean up
            workbook.Dispose();
            loadedWorkbook.Dispose();
        }
    }
}
