using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionUpgrade
{
    class Program
    {
        static void Main()
        {
            // Path to the legacy‑encrypted workbook
            string inputPath = "LegacyEncryptedWorkbook.xls";

            // Password used for the legacy encryption (if any)
            string legacyPassword = "oldPassword";

            // Load the workbook using the legacy password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = legacyPassword;
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Set a new password for the upgraded encryption
            string newPassword = "newStrongPassword";
            workbook.Settings.Password = newPassword;

            // Upgrade encryption to AES‑256 (StrongCryptographicProvider, 256‑bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the workbook with the new strong encryption
            string outputPath = "UpgradedEncryptedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Optional: verify that the workbook is now encrypted
            Workbook loaded = new Workbook(outputPath, new LoadOptions { Password = newPassword });
            Console.WriteLine("IsEncrypted after upgrade: " + loaded.Settings.IsEncrypted);
        }
    }
}