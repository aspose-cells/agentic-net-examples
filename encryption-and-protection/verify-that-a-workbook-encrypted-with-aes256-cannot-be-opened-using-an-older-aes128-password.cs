// Title: AES‑256 encrypted Excel workbook cannot be opened with an AES‑128 password – Aspose.Cells for .NET demo
// Description: Shows how to create a workbook, encrypt it with AES‑256 (256‑bit key) using Aspose.Cells, verify the file is encrypted, open it with the correct password, and confirm that loading with an incorrect password (simulating AES‑128) fails and FileFormatUtil.VerifyPassword returns false.
// Keywords: Aspose.Cells | C# | AES-256 encryption | Excel workbook protection | SetEncryptionOptions | FileFormatUtil | VerifyPassword | incorrect password handling | encryption validation | strong cryptographic provider
// Common Searches: Aspose.Cells AES-256 encryption example | verify password for encrypted Excel file .NET | open AES-256 protected workbook with wrong password | detect encrypted Excel file using Aspose.Cells | ensure AES-256 workbook cannot be opened with AES-128 password
// Developer Intent: The developer wants to confirm that a workbook encrypted with AES‑256 cannot be opened using an AES‑128 (incorrect) password.
// Use Cases: Create and save an Excel workbook encrypted with AES‑256 and check IsEncrypted via FileFormatUtil.DetectFileFormat. | Load the encrypted workbook with the correct password to verify successful decryption and that Settings.IsEncrypted becomes false. | Attempt to load the same file with an incorrect password and use FileFormatUtil.VerifyPassword to demonstrate that decryption fails and returns false.
// AI Prompts: Provide C# code that encrypts an Excel workbook with AES‑256 using Aspose.Cells, then shows the exception thrown when opening it with a wrong password and the false result from VerifyPassword. | Explain how Aspose.Cells distinguishes between AES‑256 and AES‑128 encrypted files and why an incorrect password cannot decrypt an AES‑256 workbook. | Generate a unit test in C# that asserts IsEncrypted is true after AES‑256 encryption and that opening the file with a wrong password fails both on load and VerifyPassword.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // Shows how to create a workbook, encrypt it with AES‑256 (256‑bit key) using Aspose.Cells, verify the file is encrypted, open it with the correct password, and confirm that loading with an incorrect password (simulating AES‑128) fails and FileFormatUtil.VerifyPassword returns false.
    class Program
    {
        static void Main()
        {
            // Path for the AES‑256 encrypted workbook
            string aes256Path = "aes256_encrypted.xlsx";

            // -----------------------------------------------------------------
            // Create a new workbook and encrypt it with AES‑256 (key length 256)
            // -----------------------------------------------------------------
            Workbook wb = new Workbook();                     // create workbook
            wb.Worksheets[0].Cells["A1"].PutValue("Secret Data");
            wb.Settings.Password = "SecretPwd";               // set password
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256); // AES‑256
            wb.Save(aes256Path);                              // save encrypted workbook

            // ---------------------------------------------------------------
            // Verify that the workbook is indeed encrypted (IsEncrypted = true)
            // ---------------------------------------------------------------
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(aes256Path);
            Console.WriteLine($"File encrypted? {info.IsEncrypted}"); // expected: True

            // ---------------------------------------------------------------
            // Attempt to open the AES‑256 workbook with the correct password
            // ---------------------------------------------------------------
            LoadOptions correctLoad = new LoadOptions { Password = "SecretPwd" };
            Workbook openedCorrect = new Workbook(aes256Path, correctLoad);
            Console.WriteLine($"Opened with correct password, IsEncrypted after load: {openedCorrect.Settings.IsEncrypted}"); // expected: False

            // ---------------------------------------------------------------
            // Attempt to open the same workbook with an incorrect (AES‑128) password
            // ---------------------------------------------------------------
            LoadOptions wrongLoad = new LoadOptions { Password = "WrongPwd" };
            try
            {
                Workbook openedWrong = new Workbook(aes256Path, wrongLoad);
                // If no exception, verify password using VerifyPassword (should be false)
                bool isValid = FileFormatUtil.VerifyPassword(File.OpenRead(aes256Path), "WrongPwd");
                Console.WriteLine($"Opened with wrong password, verification result: {isValid}");
            }
            catch (Exception ex)
            {
                // Expected path: loading fails due to incorrect password
                Console.WriteLine($"Failed to open with wrong password (AES‑128 simulation): {ex.Message}");
            }

            // ---------------------------------------------------------------
            // Additional check using VerifyPassword directly (should return false)
            // ---------------------------------------------------------------
            using (Stream stream = File.OpenRead(aes256Path))
            {
                bool verifyResult = FileFormatUtil.VerifyPassword(stream, "WrongPwd");
                Console.WriteLine($"VerifyPassword with wrong password returns: {verifyResult}"); // expected: False
            }
        }
    }
}
