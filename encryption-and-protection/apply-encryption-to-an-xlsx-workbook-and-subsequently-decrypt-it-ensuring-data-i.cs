using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create and encrypt workbook --------------------
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Set a password for the workbook
            string password = "StrongPassword123";
            workbook.Settings.Password = password;

            // Define encryption options (StrongCryptographicProvider with 128-bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedPath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedPath);
            Console.WriteLine($"Encrypted workbook saved to '{encryptedPath}'.");

            // -------------------- Load and verify encrypted workbook --------------------
            // Prepare load options with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            // Load the encrypted workbook using the password
            Workbook loadedEncrypted = new Workbook(encryptedPath, loadOptions);

            // Verify that the data is intact
            string cellValue = loadedEncrypted.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Loaded encrypted workbook cell A1 value: '{cellValue}'");

            // -------------------- Decrypt and save unprotected workbook --------------------
            // Remove password protection
            loadedEncrypted.Settings.Password = null;

            // Save the workbook without encryption
            string decryptedPath = "DecryptedWorkbook.xlsx";
            loadedEncrypted.Save(decryptedPath);
            Console.WriteLine($"Decrypted workbook saved to '{decryptedPath}'.");

            // -------------------- Load and verify decrypted workbook --------------------
            Workbook loadedDecrypted = new Workbook(decryptedPath);
            string decryptedCellValue = loadedDecrypted.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Loaded decrypted workbook cell A1 value: '{decryptedCellValue}'");

            // Confirm data integrity
            bool isDataIntact = cellValue == decryptedCellValue;
            Console.WriteLine($"Data integrity check passed: {isDataIntact}");
        }
    }
}