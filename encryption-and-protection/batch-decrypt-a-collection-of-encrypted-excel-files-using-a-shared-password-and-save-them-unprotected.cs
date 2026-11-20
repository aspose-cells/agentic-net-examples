// Title: Batch decrypt multiple encrypted .xlsx workbooks using a shared password with Aspose.Cells for .NET (C#)
// Common Searches: C# code to decrypt all password‑protected Excel files in a folder using Aspose.Cells | How to remove workbook encryption from many .xlsx files with a single password in .NET | Aspose.Cells example for loading encrypted Excel workbook and saving it unprotected | Automate batch decryption of encrypted Excel spreadsheets using a shared password | Process a directory of protected .xlsx files and output decrypted copies with Aspose.Cells
// Tags: batch decrypt xlsx Aspose.Cells | load encrypted workbook with password C# | remove workbook password Aspose.Cells | save unprotected Excel file C# | process multiple Excel files Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ExcelBatchProcessing
{
    // The program enumerates .xlsx files in a source folder, loads each encrypted workbook with a shared password via Aspose.Cells LoadOptions, clears the workbook's password, and saves the decrypted version to a target folder with a '_decrypted.xlsx' suffix.
    class BatchDecryptExcel
    {
        static void Main()
        {
            // Shared password for all encrypted Excel files
            const string sharedPassword = "YourSharedPassword";

            // Folder containing the encrypted Excel files
            string sourceFolder = @"C:\EncryptedExcels";

            // Folder to save the decrypted (unprotected) Excel files
            string targetFolder = @"C:\DecryptedExcels";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the target folder exists
            Directory.CreateDirectory(targetFolder);

            // Get all Excel files in the source folder
            IEnumerable<string> encryptedFiles = Directory.EnumerateFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string encryptedFilePath in encryptedFiles)
            {
                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(encryptedFilePath))
                    {
                        Console.WriteLine($"File not found (skipped): {encryptedFilePath}");
                        continue;
                    }

                    // Load the encrypted workbook using the shared password
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = sharedPassword
                    };
                    Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

                    // Remove workbook-level encryption
                    workbook.Settings.Password = null;

                    // Build the output file path
                    string fileName = Path.GetFileNameWithoutExtension(encryptedFilePath);
                    string decryptedFilePath = Path.Combine(targetFolder, $"{fileName}_decrypted.xlsx");

                    // Save the workbook without any password (unprotected)
                    workbook.Save(decryptedFilePath, SaveFormat.Xlsx);

                    Console.WriteLine($"Decrypted and saved: {decryptedFilePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{encryptedFilePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch decryption completed.");
        }
    }
}
