using System;
using System.IO;
using Aspose.Cells;

namespace BatchEncryptExcel
{
    class Program
    {
        static void Main()
        {
            // Directory containing the Excel files to encrypt
            string sourceDirectory = @"C:\InputExcelFiles";

            // Directory where encrypted files will be saved
            string destinationDirectory = @"C:\EncryptedExcelFiles";

            // Shared password for all workbooks
            string sharedPassword = "MySharedPassword";

            try
            {
                // Ensure the source directory exists
                if (!Directory.Exists(sourceDirectory))
                {
                    Console.WriteLine($"Source directory not found: {sourceDirectory}");
                    return;
                }

                // Ensure the destination directory exists
                Directory.CreateDirectory(destinationDirectory);

                // Process each Excel file in the source directory
                foreach (string filePath in Directory.GetFiles(sourceDirectory, "*.*", SearchOption.TopDirectoryOnly))
                {
                    // Consider only Excel formats (xls, xlsx, xlsb, xlsm)
                    string extension = Path.GetExtension(filePath).ToLowerInvariant();
                    if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsb" && extension != ".xlsm")
                        continue;

                    // Verify the file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook workbook = new Workbook(filePath);

                        // Set the password that protects opening the file
                        workbook.Settings.Password = sharedPassword;

                        // Apply AES‑128 encryption (StrongCryptographicProvider with 128‑bit key)
                        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                        // Save the encrypted workbook to the destination folder (overwrites if exists)
                        string destPath = Path.Combine(destinationDirectory, Path.GetFileName(filePath));
                        workbook.Save(destPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch encryption completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}