// Title: C# example: logging timestamp, file name, and encryption action for an Excel workbook with Aspose.Cells
// Common Searches: C# Aspose.Cells log encryption activity with timestamp and file name | record workbook encryption details to a text file using Aspose.Cells | how to append encryption actions to a log when protecting Excel files in C# | save encrypted Excel workbook and write audit log with Aspose.Cells | Aspose.Cells password protect workbook and track operation in log
// Tags: Aspose.Cells workbook encryption audit | timestamped log for encrypted Excel files | C# append encryption action to log | password protect Excel workbook Aspose.Cells | record file name during workbook encryption

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptionLogger
{
    // Simple logger that appends entries to a log file.
    // The sample defines a static Logger that appends a line containing the current timestamp, the workbook file name, and the performed action to a text file. It loads an Excel workbook with Aspose.Cells, sets a password to encrypt it, saves the encrypted file, and then records the encryption event using the Logger.
    public static class Logger
    {
        private static readonly string LogFilePath = "encryption_log.txt";

        // Records timestamp, file name and performed action.
        public static void Log(string fileName, string action)
        {
            string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\t{fileName}\t{action}";
            File.AppendAllLines(LogFilePath, new[] { entry });
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example workbook file to encrypt.
            string sourcePath = "Sample.xlsx";
            string encryptedPath = "Sample_Encrypted.xlsx";
            string password = "MySecretPassword";

            // Load the workbook (using the provided load rule).
            Workbook workbook = new Workbook(sourcePath);

            // Apply encryption by setting the password.
            workbook.Settings.Password = password;

            // Save the encrypted workbook (using the provided save rule).
            workbook.Save(encryptedPath, SaveFormat.Xlsx);

            // Log the encryption action.
            Logger.Log(encryptedPath, "Encrypted with password");
        }
    }
}
