using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsEncryptionUpgrade
{
    class Program
    {
        static void Main()
        {
            // Paths and passwords
            string inputFile = "legacy_encrypted.ods";   // ODS file encrypted with legacy algorithm
            string oldPassword = "oldPassword";          // Password used for the legacy encryption
            string newPassword = "newPassword";          // Password for the new AES‑256 encryption
            string outputFile = "upgraded_aes256.ods";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputFile)}");
                    return;
                }

                // Detect if the file is encrypted (optional, just for demonstration)
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputFile);
                Console.WriteLine($"Is encrypted (detected): {formatInfo.IsEncrypted}");

                // Load the legacy‑encrypted ODS file using the old password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods)
                {
                    Password = oldPassword
                };
                Workbook workbook = new Workbook(inputFile, loadOptions);

                // Apply new password and AES‑256 encryption
                workbook.Settings.Password = newPassword;
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

                // Save the workbook with the upgraded encryption
                OdsSaveOptions saveOptions = new OdsSaveOptions();
                workbook.Save(outputFile, saveOptions);

                Console.WriteLine("File decrypted and re‑encrypted with AES‑256 successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}