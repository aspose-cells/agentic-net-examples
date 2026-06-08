using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

namespace AsposeCellsEncryptionHashDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data that needs protection.");

            // Set a password to encrypt the workbook
            workbook.Settings.Password = "StrongPassword!123";

            // (Optional) Define encryption options – strong encryption with 256‑bit key
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Define the output file path
            string encryptedFilePath = "EncryptedWorkbook.xlsx";

            // Save the encrypted workbook
            workbook.Save(encryptedFilePath);
            workbook.Dispose();

            // Compute SHA‑256 hash of the saved file for integrity verification
            byte[] fileBytes = File.ReadAllBytes(encryptedFilePath);
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(fileBytes);
                string hashHex = BitConverter.ToString(hashBytes).Replace("-", string.Empty);
                Console.WriteLine($"SHA‑256 hash of the encrypted workbook: {hashHex}");
            }
        }
    }
}