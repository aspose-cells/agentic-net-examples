using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // 2. Set a password to encrypt the workbook
            workbook.Settings.Password = "Secret123";

            // 3. (Optional) Define encryption options (strong AES 128‑bit)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // 4. Save the encrypted workbook
            string encryptedPath = "encrypted.xlsx";
            workbook.Save(encryptedPath, SaveFormat.Xlsx);

            // 5. Verify that the workbook reports being encrypted
            Console.WriteLine($"Workbook IsEncrypted after save: {workbook.Settings.IsEncrypted}");

            // 6. Attempt to open the encrypted file with an incorrect password using LoadOptions
            try
            {
                LoadOptions wrongOptions = new LoadOptions { Password = "WrongPass" };
                Workbook wrongLoad = new Workbook(encryptedPath, wrongOptions);
                // If no exception, the password was somehow accepted (unexpected)
                Console.WriteLine("Unexpectedly opened workbook with wrong password.");
            }
            catch (Exception ex)
            {
                // Expected path: Aspose throws an exception for wrong password
                Console.WriteLine($"Failed to open with wrong password as expected: {ex.Message}");
            }

            // 7. Use FileFormatUtil.VerifyPassword to programmatically confirm the password is invalid
            using (Stream stream = File.OpenRead(encryptedPath))
            {
                bool isValid = FileFormatUtil.VerifyPassword(stream, "WrongPass");
                Console.WriteLine($"FileFormatUtil.VerifyPassword with wrong password returned: {isValid}");
            }

            // 8. Finally, open with the correct password to demonstrate successful decryption
            LoadOptions correctOptions = new LoadOptions { Password = "Secret123" };
            Workbook correctLoad = new Workbook(encryptedPath, correctOptions);
            Console.WriteLine($"Successfully opened with correct password. Cell A1 value: {correctLoad.Worksheets[0].Cells["A1"].Value}");
        }
    }
}