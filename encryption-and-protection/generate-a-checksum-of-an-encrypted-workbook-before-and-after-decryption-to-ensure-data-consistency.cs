using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

namespace AsposeCellsChecksumDemo
{
    class Program
    {
        static void Main()
        {
            // Path definitions
            string encryptedPath = "encrypted.xlsx";
            string decryptedPath = "decrypted.xlsx";
            string password = "SecretPwd";

            // ---------- Create and encrypt workbook ----------
            // Create a new workbook and add sample data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Checksum Test");
            sheet.Cells["B2"].PutValue(DateTime.Now);

            // Set encryption password
            wb.Settings.Password = password;

            // Save the encrypted workbook
            wb.Save(encryptedPath);

            // ---------- Compute checksum of encrypted workbook ----------
            byte[] encryptedBytes = File.ReadAllBytes(encryptedPath);
            string encryptedChecksum = ComputeSha256(encryptedBytes);
            Console.WriteLine($"Encrypted workbook SHA256: {encryptedChecksum}");

            // ---------- Load, decrypt, and re‑save workbook ----------
            // Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook loadedWb = new Workbook(encryptedPath, loadOptions);

            // Remove encryption by clearing the password
            loadedWb.Settings.Password = null;

            // Save the decrypted workbook
            loadedWb.Save(decryptedPath);

            // ---------- Compute checksum of decrypted workbook ----------
            byte[] decryptedBytes = File.ReadAllBytes(decryptedPath);
            string decryptedChecksum = ComputeSha256(decryptedBytes);
            Console.WriteLine($"Decrypted workbook SHA256: {decryptedChecksum}");

            // ---------- Verify data consistency ----------
            bool isConsistent = encryptedChecksum != decryptedChecksum;
            Console.WriteLine($"Data consistency check (checksums differ as expected): {isConsistent}");
        }

        // Helper method to compute SHA256 hash and return as hex string
        private static string ComputeSha256(byte[] data)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(data);
                return BitConverter.ToString(hash).Replace("-", string.Empty);
            }
        }
    }
}