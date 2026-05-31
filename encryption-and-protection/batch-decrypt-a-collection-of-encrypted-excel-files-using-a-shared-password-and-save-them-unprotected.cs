using System;
using System.IO;
using Aspose.Cells;

namespace BatchDecryptExcel
{
    class Program
    {
        static void Main()
        {
            // Folder containing encrypted Excel files
            string sourceFolder = @"C:\EncryptedFiles";
            // Folder where unprotected files will be saved
            string destinationFolder = @"C:\DecryptedFiles";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the destination folder exists
            Directory.CreateDirectory(destinationFolder);

            // Shared password for all encrypted workbooks
            string sharedPassword = "MySharedPassword";

            // Process each file in the source folder
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly))
            {
                // Skip non‑Excel files based on extension (optional)
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".ods")
                    continue;

                // Verify the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook using the shared password
                    LoadOptions loadOptions = new LoadOptions { Password = sharedPassword };
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // Remove workbook structure/window protection if present
                    try
                    {
                        workbook.Unprotect(sharedPassword);
                    }
                    catch
                    {
                        // Ignore if workbook is not protected
                    }

                    // Remove file‑level encryption
                    workbook.Settings.Password = null;

                    // Build the output file path
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    string outputPath = Path.Combine(destinationFolder, $"{fileName}_decrypted{ext}");

                    // Save the unprotected workbook
                    workbook.Save(outputPath);
                    Console.WriteLine($"Decrypted: {Path.GetFileName(filePath)} → {Path.GetFileName(outputPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch decryption completed.");
        }
    }
}