// Title: Encrypt an Excel workbook with a password using Aspose.Cells .NET and confirm decryption fails with a wrong password
// Description: This C# example creates a workbook, writes data to cell A1, applies a password and strong 128‑bit encryption, saves the file, checks the IsEncrypted flag, uses FileFormatUtil.VerifyPassword with an incorrect password, attempts to load the file with a wrong password to capture the expected exception, and finally opens it with the correct password to demonstrate successful decryption.
// Keywords: Aspose.Cells encrypt workbook | Excel password protection .NET | SetEncryptionOptions Aspose | FileFormatUtil VerifyPassword example | load encrypted workbook wrong password | Workbook.IsEncrypted property | C# Excel encryption Aspose
// Common Searches: how to password‑protect an xlsx file with Aspose.Cells | check if an Excel file is encrypted using Aspose | verify decryption failure with wrong password Aspose.Cells | C# example for workbook encryption and error handling | Aspose.Cells strong encryption options
// Developer Intent: Apply password protection to an Excel workbook, ensure the file reports as encrypted, and validate that opening it with an incorrect password throws an error.
// Use Cases: Secure confidential spreadsheets before distribution. | Test that external systems cannot bypass password protection. | Implement robust error handling for encrypted workbook loading in automated pipelines.
// AI Prompts: Generate C# code that encrypts an Excel workbook with 256‑bit AES using Aspose.Cells and verifies the IsEncrypted flag after saving. | Show how to catch the specific Aspose.Cells exception when loading an encrypted workbook with an invalid password.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // This C# example creates a workbook, writes data to cell A1, applies a password and strong 128‑bit encryption, saves the file, checks the IsEncrypted flag, uses FileFormatUtil.VerifyPassword with an incorrect password, attempts to load the file with a wrong password to capture the expected exception, and finally opens it with the correct password to demonstrate successful decryption.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // 2. Set a password to encrypt the workbook
            workbook.Settings.Password = "Secret123";

            // Optional: define encryption options (e.g., strong encryption, 128-bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // 3. Save the encrypted workbook
            string encryptedPath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedPath, SaveFormat.Xlsx);

            // Verify that the workbook reports being encrypted
            Console.WriteLine($"Workbook.IsEncrypted after save: {workbook.Settings.IsEncrypted}");

            // ------------------------------------------------------------
            // 4. Attempt decryption with an incorrect password using Aspose's
            //    verification method (simulating a third‑party check)
            // ------------------------------------------------------------
            bool isPasswordCorrect = FileFormatUtil.VerifyPassword(
                File.OpenRead(encryptedPath), "WrongPassword");
            Console.WriteLine($"FileFormatUtil.VerifyPassword with wrong password: {isPasswordCorrect}");

            // ------------------------------------------------------------
            // 5. Attempt to load the encrypted workbook with a wrong password
            //    and catch the expected exception
            // ------------------------------------------------------------
            try
            {
                LoadOptions loadOptions = new LoadOptions { Password = "WrongPassword" };
                Workbook wrongLoad = new Workbook(encryptedPath, loadOptions);
                // If no exception, the decryption unexpectedly succeeded
                Console.WriteLine("Unexpectedly opened workbook with wrong password.");
            }
            catch (Exception ex)
            {
                // Expected path: decryption fails
                Console.WriteLine($"Failed to open workbook with wrong password: {ex.Message}");
            }

            // ------------------------------------------------------------
            // 6. Load the workbook with the correct password to demonstrate success
            // ------------------------------------------------------------
            LoadOptions correctLoadOptions = new LoadOptions { Password = "Secret123" };
            Workbook correctLoad = new Workbook(encryptedPath, correctLoadOptions);
            Console.WriteLine($"Successfully opened workbook. Cell A1 value: {correctLoad.Worksheets[0].Cells["A1"].Value}");
        }
    }
}
