// Title: Encrypt an Excel workbook with Aspose.Cells, test AES decryption failure, and verify password
// Description: Creates a workbook, writes sensitive data, applies a password using Aspose.Cells (optionally with strong encryption), saves it as XLSX, then attempts a naive AES decryption that inevitably fails, and finally validates the password with FileFormatUtil.VerifyPassword for both wrong and correct keys.
// Keywords: Aspose.Cells encrypt workbook | C# Excel password protection | FileFormatUtil VerifyPassword | AES decryption attempt on XLSX | strong encryption options Aspose | OOXML password verification | third‑party decryption test
// Common Searches: How to password‑protect an XLSX file using Aspose.Cells .NET | Verify Excel file password without opening the workbook | Can a custom AES routine decrypt an Aspose‑encrypted workbook | Set strong encryption options for Excel 2003 format in Aspose.Cells | Example of failed decryption of an encrypted XLSX file
// Developer Intent: Show how to encrypt an Excel file with Aspose.Cells, demonstrate that generic AES code cannot decrypt it, and confirm password validity via Aspose's verification API.
// Use Cases: Secure distribution of confidential spreadsheets and ensure they resist simple AES attacks. | Programmatically validate a user‑supplied password before loading the workbook. | Compliance testing to prove that third‑party libraries cannot bypass Aspose.Cells encryption.
// AI Prompts: Generate C# code that encrypts an XLSX workbook with Aspose.Cells using a strong cryptographic provider and then checks the password with FileFormatUtil. | Write a method that tries to decrypt an Aspose‑encrypted XLSX file using AES and returns false on failure, handling all exceptions. | Explain how to test that an incorrect password is rejected for a workbook encrypted with Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // Creates a workbook, writes sensitive data, applies a password using Aspose.Cells (optionally with strong encryption), saves it as XLSX, then attempts a naive AES decryption that inevitably fails, and finally validates the password with FileFormatUtil.VerifyPassword for both wrong and correct keys.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Set a password to encrypt the workbook
            wb.Settings.Password = "Secret123";

            // Optionally set stronger encryption options (Excel 2003 specific, ignored for newer formats)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedPath = "encrypted.xlsx";
            wb.Save(encryptedPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved encrypted at '{encryptedPath}'.");

            // -----------------------------------------------------------------
            // Attempt to decrypt the file using a third‑party library (AES demo)
            // This will fail because the file is not a plain AES encrypted blob.
            // -----------------------------------------------------------------
            bool thirdPartyDecryptionResult = TryDecryptWithAes(encryptedPath, "WrongPass");
            Console.WriteLine($"Third‑party AES decryption succeeded? {thirdPartyDecryptionResult}");

            // Verify failure using Aspose's built‑in password verification (wrong password)
            bool asposePasswordCheck = FileFormatUtil.VerifyPassword(File.OpenRead(encryptedPath), "WrongPass");
            Console.WriteLine($"Aspose password verification with wrong password: {asposePasswordCheck}");

            // Verify success with correct password
            bool asposeCorrectCheck = FileFormatUtil.VerifyPassword(File.OpenRead(encryptedPath), "Secret123");
            Console.WriteLine($"Aspose password verification with correct password: {asposeCorrectCheck}");
        }

        // Demonstrates a naive AES decryption attempt on the encrypted OOXML file.
        // Returns true only if decryption completes without exception (which it shouldn't).
        static bool TryDecryptWithAes(string filePath, string password)
        {
            try
            {
                byte[] fileBytes = File.ReadAllBytes(filePath);

                // Derive a key and IV from the password (for demonstration only)
                var pdb = new Rfc2898DeriveBytes(password, Encoding.UTF8.GetBytes("salt1234"));
                byte[] key = pdb.GetBytes(32); // 256‑bit key
                byte[] iv = pdb.GetBytes(16);  // 128‑bit IV

                using (Aes aes = Aes.Create())
                {
                    aes.Key = key;
                    aes.IV = iv;

                    using (var ms = new MemoryStream())
                    using (var cs = new CryptoStream(ms, aes.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(fileBytes, 0, fileBytes.Length);
                        cs.FlushFinalBlock(); // Will throw if padding is invalid
                    }
                }

                // If we reach here, decryption (incorrectly) succeeded
                return true;
            }
            catch
            {
                // Expected path: decryption fails
                return false;
            }
        }
    }
}
