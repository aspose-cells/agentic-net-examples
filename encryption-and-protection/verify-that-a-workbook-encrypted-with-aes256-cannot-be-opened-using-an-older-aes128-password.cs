using System;
using System.IO;
using Aspose.Cells;

class VerifyAes256Encryption
{
    static void Main()
    {
        // Paths for the temporary workbook
        string encryptedPath = "Aes256Encrypted.xlsx";

        // ---------- Create and encrypt workbook with AES‑256 ----------
        Workbook wb = new Workbook();                         // create
        wb.Worksheets[0].Cells["A1"].PutValue("Secret Data");
        wb.Settings.Password = "StrongPassword";              // set password
        // Set encryption to AES‑256 (key length 256 bits)
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);
        wb.Save(encryptedPath);                               // save

        // ---------- Attempt to open with the correct password ----------
        try
        {
            LoadOptions correctLoad = new LoadOptions { Password = "StrongPassword" };
            Workbook openedCorrect = new Workbook(encryptedPath, correctLoad); // load
            Console.WriteLine("Opened with correct password: " + openedCorrect.Worksheets[0].Cells["A1"].Value);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to open with correct password: " + ex.Message);
        }

        // ---------- Attempt to open with an older AES‑128 password (simulated by wrong password) ----------
        try
        {
            LoadOptions wrongLoad = new LoadOptions { Password = "OldPassword128" };
            Workbook openedWrong = new Workbook(encryptedPath, wrongLoad); // load
            Console.WriteLine("Unexpectedly opened with old AES‑128 password.");
        }
        catch (Exception ex)
        {
            // Expected failure because the workbook is encrypted with AES‑256
            Console.WriteLine("Failed to open with old AES‑128 password (as expected): " + ex.Message);
        }

        // Clean up temporary file
        if (File.Exists(encryptedPath))
            File.Delete(encryptedPath);
    }
}