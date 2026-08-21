// Title: Compute SHA256 Checksums of Aspose.Cells Workbooks Before Encryption and After Decryption (C#)
// Description: This example creates an Excel workbook with Aspose.Cells, calculates a SHA256 hash of the unencrypted file, encrypts it with a password, hashes the encrypted file, then loads and decrypts the workbook, re‑hashes the result, and compares the two checksums to confirm data integrity.
// Keywords: Aspose.Cells | C# | .NET | SHA256 checksum | workbook integrity | Excel encryption | password‑protected .xlsx | LoadOptions | SaveFormat | data consistency verification | memory stream hash
// Common Searches: Aspose.Cells compute SHA256 hash of workbook C# | verify Excel file integrity after decryption Aspose | checksum encrypted .xlsx using Aspose.Cells | compare original and decrypted workbook hashes | C# example for password‑protected Excel checksum
// Developer Intent: Generate a SHA256 hash of a workbook before it is encrypted, generate a second hash after decryption, and compare the two values to ensure the content has not changed.
// Use Cases: Confirm that password‑protected Excel reports can be restored without data loss. | Detect tampering or corruption of encrypted workbooks by comparing pre‑ and post‑decryption hashes. | Automate integrity validation in CI/CD pipelines for Excel files generated with Aspose.Cells.
// AI Prompts: Show C# code that uses Aspose.Cells to calculate a SHA256 checksum of a workbook saved to a MemoryStream, then verify it after loading with a password. | Explain why clearing the workbook password before re‑saving is required for matching checksums. | Provide a pattern for batch processing multiple workbooks, each encrypted with a different password, and validating their checksums.

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

namespace AsposeCellsChecksumDemo
{
    // This example creates an Excel workbook with Aspose.Cells, calculates a SHA256 hash of the unencrypted file, encrypts it with a password, hashes the encrypted file, then loads and decrypts the workbook, re‑hashes the result, and compares the two checksums to confirm data integrity.
    class Program
    {
        // Compute SHA256 checksum of a byte array and return as hex string
        static string ComputeChecksum(byte[] data)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(data);
                return BitConverter.ToString(hash).Replace("-", string.Empty);
            }
        }

        static void Main()
        {
            // ---------- Create a new workbook and add sample data ----------
            Workbook originalWorkbook = new Workbook();
            Worksheet sheet = originalWorkbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Checksum Test");
            sheet.Cells["B2"].PutValue(12345);
            sheet.Cells["C3"].PutValue(DateTime.Now);

            // ---------- Save original workbook to memory (unencrypted) ----------
            byte[] originalBytes;
            using (MemoryStream ms = new MemoryStream())
            {
                originalWorkbook.Save(ms, SaveFormat.Xlsx);
                originalBytes = ms.ToArray();
            }
            string originalChecksum = ComputeChecksum(originalBytes);
            Console.WriteLine($"Original (unencrypted) checksum: {originalChecksum}");

            // ---------- Encrypt the workbook with a password ----------
            string password = "SecretPwd123";
            originalWorkbook.Settings.Password = password;
            string encryptedPath = "encrypted_workbook.xlsx";
            originalWorkbook.Save(encryptedPath); // saved encrypted

            // ---------- Compute checksum of the encrypted file ----------
            byte[] encryptedBytes = File.ReadAllBytes(encryptedPath);
            string encryptedChecksum = ComputeChecksum(encryptedBytes);
            Console.WriteLine($"Encrypted file checksum: {encryptedChecksum}");

            // ---------- Load the encrypted workbook using the password ----------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            loadOptions.Password = password;
            Workbook decryptedWorkbook = new Workbook(encryptedPath, loadOptions);

            // ---------- Save the decrypted workbook to memory (without password) ----------
            byte[] decryptedBytes;
            using (MemoryStream ms = new MemoryStream())
            {
                // Ensure password is cleared before saving
                decryptedWorkbook.Settings.Password = null;
                decryptedWorkbook.Save(ms, SaveFormat.Xlsx);
                decryptedBytes = ms.ToArray();
            }
            string decryptedChecksum = ComputeChecksum(decryptedBytes);
            Console.WriteLine($"Decrypted (after loading) checksum: {decryptedChecksum}");

            // ---------- Verify data consistency ----------
            bool isConsistent = originalChecksum.Equals(decryptedChecksum, StringComparison.OrdinalIgnoreCase);
            Console.WriteLine($"Data consistency check passed: {isConsistent}");
        }
    }
}
