using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionUpgrade
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source Excel file (could be legacy .xls with old encryption)
                string sourcePath = "legacy_encrypted.xls";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Password required to open the encrypted file
                string password = "oldPassword";

                // Detect file format and encryption status
                FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(sourcePath);
                Console.WriteLine($"Is Encrypted: {fileInfo.IsEncrypted}");
                Console.WriteLine($"File Format Type: {fileInfo.FileFormatType}");

                // If the file is encrypted and not already in a modern format (e.g., XLSX), upgrade it
                if (fileInfo.IsEncrypted && fileInfo.FileFormatType != FileFormatType.Xlsx)
                {
                    // Load the encrypted workbook using the provided password
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                    {
                        Password = password
                    };
                    Workbook workbook = new Workbook(sourcePath, loadOptions);
                    Console.WriteLine("Legacy encrypted workbook loaded successfully.");

                    // Apply modern encryption settings (Strong Cryptographic Provider, 128‑bit key)
                    workbook.Settings.Password = password; // keep the same password
                    workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                    // Save the workbook in a modern format (XLSX) with the new encryption
                    string upgradedPath = "upgraded.xlsx";
                    workbook.Save(upgradedPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook upgraded and saved to '{upgradedPath}'.");
                }
                else
                {
                    Console.WriteLine("File is either not encrypted or already uses a modern format. No upgrade needed.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}