// Title: How to encrypt an Excel workbook with a password and record timestamp, file name, and action in a log using Aspose.Cells for .NET
// Common Searches: asp.net encrypt excel file with password using aspose.cells and write log entry | c# log workbook encryption timestamp and filename after saving with Aspose.Cells | append encryption activity to a text file after saving encrypted xlsx in .NET | record password protection action for an Excel workbook in a custom log using Aspose.Cells | create audit log for encrypted Excel files in C#
// Tags: Aspose.Cells XLSX password encryption | C# Excel workbook encryption logging | timestamped log entry for encrypted workbooks | append audit log after saving Excel file | record encryption activity in .NET

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptionLogger
{
    // The example loads or creates an Excel workbook, applies a password via Aspose.Cells Settings.Password, saves the file as an encrypted XLSX, and then appends a log line containing the current timestamp, output file name, and the encryption action to a text file, with error handling that writes exceptions to the console.
    public static class EncryptionProcessor
    {
        // Path to the log file (can be changed as needed)
        private const string LogFilePath = "encryption_log.txt";

        public static void EncryptAndLog(string inputPath, string outputPath, string password)
        {
            try
            {
                Workbook workbook;

                // Load existing workbook or create a new one if the input file does not exist
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");
                    workbook.Save(inputPath);
                }

                // Set password for opening the workbook (encryption)
                workbook.Settings.Password = password;

                // Save the encrypted workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                // Build log entry
                string fileName = Path.GetFileName(outputPath);
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string logEntry = $"{timestamp} | File: {fileName} | Action: Encrypted with password";

                // Append log entry to the log file
                File.AppendAllText(LogFilePath, logEntry + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // Output error information for troubleshooting
                Console.Error.WriteLine($"Error during encryption: {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string sourceFile = "input.xlsx";
            string encryptedFile = "input_encrypted.xlsx";
            string password = "MySecretPassword";

            EncryptionProcessor.EncryptAndLog(sourceFile, encryptedFile, password);
        }
    }
}
