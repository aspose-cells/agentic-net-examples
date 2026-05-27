using System;
using System.IO;
using Aspose.Cells;

class VerifyAes256Encryption
{
    static void Main()
    {
        try
        {
            // Password to protect the workbook
            string password = "SecretPwd";

            // Path for the encrypted workbook
            string filePath = "aes256_encrypted.xlsx";

            // -------------------------------------------------
            // Create a new workbook and add some data
            // -------------------------------------------------
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue("Encrypted with AES-256");

            // -------------------------------------------------
            // Apply password protection
            // -------------------------------------------------
            wb.Settings.Password = password;

            // -------------------------------------------------
            // Set encryption options to AES‑256 (key length = 256 bits)
            // -------------------------------------------------
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // -------------------------------------------------
            // Save the workbook (it will be encrypted with AES‑256)
            // -------------------------------------------------
            wb.Save(filePath);

            // -------------------------------------------------
            // 1) Load the workbook using the correct AES‑256 settings
            // -------------------------------------------------
            if (File.Exists(filePath))
            {
                LoadOptions loCorrect = new LoadOptions
                {
                    Password = password // correct password
                };
                Workbook wbLoadedCorrect = new Workbook(filePath, loCorrect);
                Console.WriteLine("Opened with correct AES‑256 settings: " + wbLoadedCorrect.Settings.IsEncrypted);
            }
            else
            {
                Console.WriteLine("File not found: " + filePath);
            }

            // -------------------------------------------------
            // 2) Attempt to load the same file with an incorrect password
            //    (Aspose.Cells does not expose a direct API to force AES‑128 decryption)
            // -------------------------------------------------
            if (File.Exists(filePath))
            {
                LoadOptions loWrong = new LoadOptions
                {
                    Password = "WrongPassword" // intentionally incorrect
                };

                try
                {
                    Workbook wbLoadedWrong = new Workbook(filePath, loWrong);
                    // If no exception is thrown, the test has failed
                    Console.WriteLine("Unexpectedly opened with incorrect password.");
                }
                catch (Exception ex)
                {
                    // Expected path: loading fails because the password is wrong / encryption mismatch
                    Console.WriteLine("Failed to open with incorrect password as expected: " + ex.Message);
                }
            }
            else
            {
                Console.WriteLine("File not found: " + filePath);
            }
        }
        catch (Exception e)
        {
            // Catch any unexpected runtime errors
            Console.WriteLine("An error occurred: " + e.Message);
        }
    }
}