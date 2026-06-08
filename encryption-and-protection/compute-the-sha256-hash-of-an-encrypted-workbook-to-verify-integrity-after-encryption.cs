using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHashDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for hashing");

            // Encrypt the workbook with a password
            workbook.Settings.Password = "StrongPassword123";

            // Save the encrypted workbook to a memory stream
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                byte[] encryptedBytes = ms.ToArray();

                // Compute SHA-256 hash of the encrypted workbook bytes
                byte[] hashBytes;
                using (SHA256 sha256 = SHA256.Create())
                {
                    hashBytes = sha256.ComputeHash(encryptedBytes);
                }

                // Convert hash to hexadecimal string
                StringBuilder sb = new StringBuilder();
                foreach (byte b in hashBytes)
                {
                    sb.Append(b.ToString("x2"));
                }
                string hashHex = sb.ToString();

                // Output results
                Console.WriteLine($"Workbook is encrypted: {workbook.Settings.IsEncrypted}");
                Console.WriteLine($"SHA-256 hash of encrypted workbook: {hashHex}");

                // Optionally, save the encrypted file to disk for further inspection
                File.WriteAllBytes("EncryptedWorkbook.xlsx", encryptedBytes);
            }

            // Clean up
            workbook.Dispose();
        }
    }
}