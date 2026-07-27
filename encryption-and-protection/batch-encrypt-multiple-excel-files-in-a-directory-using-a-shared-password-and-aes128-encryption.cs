using System;
using System.IO;
using Aspose.Cells;

class BatchEncrypt
{
    static void Main()
    {
        // Directory containing the Excel files to encrypt
        string folderPath = @"C:\ExcelFiles";

        // Shared password for all workbooks
        string password = "SharedPassword123";

        // Retrieve all Excel files (common extensions) in the directory
        string[] files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in files)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".xlsb")
                continue; // Skip non‑Excel files

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Set the password required to open the workbook
            workbook.Settings.Password = password;

            // Apply AES‑128 encryption (StrongCryptographicProvider with 128‑bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook, overwriting the original file
            workbook.Save(filePath);
        }
    }
}