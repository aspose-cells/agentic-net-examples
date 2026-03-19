using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    public class WorkbookEncryptionHandler
    {
        // Encrypts a workbook with a password and optional encryption options, then saves it.
        public static void EncryptWorkbook(string outputPath, string password)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Set the password that protects opening the file
            workbook.Settings.Password = password;

            // Optionally set stronger encryption (SHA AES) – Excel 2007+ ignores EncryptionType
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Protect the workbook structure as an additional layer (optional)
            workbook.Protect(ProtectionType.All, password);

            // Save the encrypted workbook (lifecycle: save)
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }

        // Decrypts a password‑protected workbook and saves an unprotected copy.
        public static void DecryptWorkbook(string encryptedPath, string password, string decryptedOutputPath)
        {
            // Load the encrypted workbook using the password (lifecycle: load)
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(encryptedPath, loadOptions);

            // Remove workbook protection if it was set
            if (workbook.IsWorkbookProtectedWithPassword)
            {
                workbook.Unprotect(password);
            }

            // Clear the opening password to make the saved file unencrypted
            workbook.Settings.Password = null;

            // Save the decrypted (unprotected) workbook
            workbook.Save(decryptedOutputPath, SaveFormat.Xlsx);
        }

        // Example usage
        public static void Main()
        {
            string encryptedFile = "EncryptedWorkbook.xlsx";
            string decryptedFile = "DecryptedWorkbook.xlsx";
            string pwd = "StrongPassword!123";

            // Encrypt the workbook
            EncryptWorkbook(encryptedFile, pwd);
            Console.WriteLine($"Workbook encrypted and saved to '{encryptedFile}'.");

            // Decrypt the workbook
            DecryptWorkbook(encryptedFile, pwd, decryptedFile);
            Console.WriteLine($"Workbook decrypted and saved to '{decryptedFile}'.");
        }
    }
}