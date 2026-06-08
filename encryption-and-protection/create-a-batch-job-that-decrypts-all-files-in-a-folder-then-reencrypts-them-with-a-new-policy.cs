using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace BatchEncryption
{
    class Program
    {
        static void Main()
        {
            // Folder containing the source Excel files
            string sourceFolder = @"C:\Input";
            // Folder where the re‑encrypted files will be saved
            string destinationFolder = @"C:\Output";

            // Old password used for decrypting existing files (if any)
            string oldPassword = "oldPassword";
            // New password to apply
            string newPassword = "newPassword";

            try
            {
                // Verify source folder exists
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder not found: {sourceFolder}");
                    return;
                }

                // Ensure the destination folder exists
                Directory.CreateDirectory(destinationFolder);

                // Process each .xlsx file in the source folder
                foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
                {
                    try
                    {
                        // Verify the file still exists
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found (skipped): {filePath}");
                            continue;
                        }

                        // Detect whether the file is encrypted
                        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                        LoadOptions loadOptions = new LoadOptions();

                        if (formatInfo.IsEncrypted)
                        {
                            // Provide the old password to open the encrypted workbook
                            loadOptions.Password = oldPassword;
                        }

                        // Load the workbook (decrypted in memory)
                        Workbook workbook = new Workbook(filePath, loadOptions);

                        // Remove any existing password/encryption
                        workbook.Settings.Password = null;

                        // Apply the new encryption policy
                        workbook.Settings.Password = newPassword;
                        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                        // Save the re‑encrypted workbook to the destination folder
                        string destPath = Path.Combine(destinationFolder, Path.GetFileName(filePath));
                        workbook.Save(destPath, SaveFormat.Xlsx);

                        Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch decryption and re‑encryption completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}