// Title: C# – Encrypt an Aspose.Cells workbook with 256‑bit AES and generate its SHA‑256 hash
// Description: This example demonstrates how to create an Excel file using Aspose.Cells, protect it with a password‑based 256‑bit AES encryption, save the encrypted workbook, and then compute a SHA‑256 checksum of the saved file for tamper detection. The resulting hash is printed as a hexadecimal string.
// Keywords: Aspose.Cells | C# encryption | Excel password protection | AES 256 | SHA-256 checksum | .NET workbook security | file integrity verification | Excel hash | Encrypt workbook C# | Compute file hash
// Common Searches: How to apply password protection to an Excel file using Aspose.Cells C# | Set 256‑bit AES encryption for a workbook in .NET | Calculate SHA‑256 hash of an Excel document after saving | Aspose.Cells encrypt workbook and verify integrity | C# code to hash encrypted Excel file
// Developer Intent: Add strong encryption to an Excel workbook and obtain a SHA‑256 digest to confirm the file has not been altered.
// Use Cases: Secure confidential reports before emailing them, then store the checksum to detect any post‑delivery modifications. | Automate archival of generated spreadsheets with password protection and record their hash for future compliance audits. | Embed encryption and hash generation into a build pipeline so released Excel artifacts are both encrypted and verifiable.
// AI Prompts: Modify the sample to use a custom encryption key instead of a password with Aspose.Cells. | Create a reusable method that returns the SHA‑256 hash as a byte array after encrypting the workbook, including proper exception handling. | Demonstrate how to recompute the hash after decrypting the file to validate its integrity.

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example demonstrates how to create an Excel file using Aspose.Cells, protect it with a password‑based 256‑bit AES encryption, save the encrypted workbook, and then compute a SHA‑256 checksum of the saved file for tamper detection. The resulting hash is printed as a hexadecimal string.
    public class WorkbookEncryptionAndHashDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle: create)
            using (Workbook workbook = new Workbook())
            {
                // Add sample data to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sensitive data that needs protection.");

                // Set a password to encrypt the workbook
                workbook.Settings.Password = "StrongPassword!123";

                // Optionally, set stronger encryption options (e.g., 256-bit AES)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

                // Define the output file path
                string encryptedFilePath = "EncryptedWorkbook.xlsx";

                // Save the encrypted workbook (lifecycle: save)
                workbook.Save(encryptedFilePath);
            }

            // Verify that the file was created before computing its hash
            string filePath = "EncryptedWorkbook.xlsx";
            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine("Encrypted file not found.");
                return;
            }

            // Compute SHA-256 hash of the saved encrypted file for integrity verification
            byte[] fileBytes = File.ReadAllBytes(filePath);
            byte[] hashBytes;
            using (SHA256 sha256 = SHA256.Create())
            {
                hashBytes = sha256.ComputeHash(fileBytes);
            }

            // Convert hash bytes to a hexadecimal string
            string hashString = BitConverter.ToString(hashBytes).Replace("-", string.Empty);

            // Output the hash
            Console.WriteLine($"SHA-256 hash of the encrypted workbook: {hashString}");
        }
    }
}
