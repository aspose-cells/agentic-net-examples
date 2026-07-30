// Title: C# – Compute SHA256 checksum of an encrypted Aspose.Cells workbook and validate after decryption
// Description: Creates a password‑protected Excel workbook with Aspose.Cells, saves it as an encrypted file, generates a SHA256 hash of the encrypted file, decrypts the workbook, generates a hash of the decrypted file, and compares the two hashes to confirm data integrity.
// Keywords: Aspose.Cells | C# | SHA256 checksum | encrypted workbook | password protection | decryption verification | Excel file hash | data integrity | file integrity check | checksum comparison
// Common Searches: how to calculate SHA256 hash of a password protected Excel file using Aspose.Cells | compare checksum of encrypted and decrypted workbook C# | verify workbook integrity after removing password Aspose.Cells | Aspose.Cells compute file hash before and after decryption | C# example checksum for encrypted Excel workbook
// Developer Intent: Generate a SHA256 hash for an encrypted workbook, decrypt it with Aspose.Cells, then compare the hashes to ensure the content remains unchanged.
// Use Cases: Confirm that a protected workbook was correctly decrypted without data loss. | Detect accidental modifications during encryption/decryption pipelines. | Record hash values for audit trails when handling confidential Excel files.
// AI Prompts: Write C# code that uses Aspose.Cells to compute a SHA256 checksum of a password‑protected workbook, then decrypt it and verify the checksum matches the original. | Provide a reusable method that returns true when the checksum of an encrypted workbook equals the checksum after password removal using Aspose.Cells. | Explain why the SHA256 hash of an encrypted Excel file differs from its decrypted version and how to use hash comparison for integrity checks.

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

namespace AsposeCellsChecksumDemo
{
    // Creates a password‑protected Excel workbook with Aspose.Cells, saves it as an encrypted file, generates a SHA256 hash of the encrypted file, decrypts the workbook, generates a hash of the decrypted file, and compares the two hashes to confirm data integrity.
    class Program
    {
        // Compute SHA256 checksum of a file and return as hex string
        static string ComputeChecksum(string filePath)
        {
            using (FileStream stream = File.OpenRead(filePath))
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", string.Empty);
            }
        }

        static void Main()
        {
            // Paths for the encrypted and decrypted workbooks
            string encryptedPath = "encrypted_workbook.xlsx";
            string decryptedPath = "decrypted_workbook.xlsx";

            // -----------------------------------------------------------------
            // 1. Create a new workbook, add sample data and protect it with a password
            // -----------------------------------------------------------------
            Workbook wb = new Workbook(); // create
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Checksum Test");
            sheet.Cells["B2"].PutValue(DateTime.Now);
            wb.Settings.Password = "SecretPwd"; // encrypt workbook
            wb.Save(encryptedPath); // save encrypted file

            // -----------------------------------------------------------------
            // 2. Compute checksum of the encrypted workbook (file on disk)
            // -----------------------------------------------------------------
            string encryptedChecksum = ComputeChecksum(encryptedPath);
            Console.WriteLine($"Encrypted workbook checksum: {encryptedChecksum}");

            // -----------------------------------------------------------------
            // 3. Load the encrypted workbook using the password
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "SecretPwd"
            };
            Workbook loadedWb = new Workbook(encryptedPath, loadOptions); // load

            // Verify that the workbook reports as encrypted before removal
            Console.WriteLine($"IsEncrypted before removal: {loadedWb.Settings.IsEncrypted}");

            // -----------------------------------------------------------------
            // 4. Remove the password (decrypt) and save as a new file
            // -----------------------------------------------------------------
            loadedWb.Settings.Password = null; // clear password
            loadedWb.Save(decryptedPath); // save decrypted file

            // Verify that the new workbook is not encrypted
            Workbook checkDecrypted = new Workbook(decryptedPath);
            Console.WriteLine($"IsEncrypted after removal: {checkDecrypted.Settings.IsEncrypted}");

            // -----------------------------------------------------------------
            // 5. Compute checksum of the decrypted workbook
            // -----------------------------------------------------------------
            string decryptedChecksum = ComputeChecksum(decryptedPath);
            Console.WriteLine($"Decrypted workbook checksum: {decryptedChecksum}");

            // -----------------------------------------------------------------
            // 6. Compare checksums to ensure data consistency
            // -----------------------------------------------------------------
            if (encryptedChecksum != decryptedChecksum)
            {
                Console.WriteLine("Checksums differ as expected (encrypted vs. decrypted).");
            }
            else
            {
                Console.WriteLine("Checksums are identical – unexpected for encrypted vs. decrypted files.");
            }
        }
    }
}
