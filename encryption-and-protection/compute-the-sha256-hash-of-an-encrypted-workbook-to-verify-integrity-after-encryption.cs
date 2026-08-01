// Title: C# – Compute SHA‑256 Hash of an Encrypted Aspose.Cells Workbook
// Description: Demonstrates how to create a workbook with Aspose.Cells, protect it with a password, save the encrypted file, calculate its SHA‑256 checksum using System.Security.Cryptography, and optionally reload the workbook with LoadOptions to confirm accessibility.
// Keywords: Aspose.Cells | C# | SHA-256 | hash encrypted workbook | password‑protected Excel | checksum verification | integrity check | LoadOptions | System.Security.Cryptography | XLSX encryption
// Common Searches: compute sha256 of password protected xlsx c# | aspocells verify encrypted workbook integrity | hash encrypted excel file using .net | aspocells loadoptions password example | c# generate checksum for encrypted workbook
// Developer Intent: Generate a SHA‑256 checksum for a password‑protected workbook to ensure its integrity.
// Use Cases: Create and encrypt an Excel file, then store its hash for tamper detection. | Validate that an encrypted workbook can still be opened after hashing. | Compare a previously saved hash with a newly computed one to detect unauthorized changes.
// AI Prompts: Write C# code that builds an Aspose.Cells workbook, encrypts it with a password, saves it, and returns the SHA‑256 hash as a hex string. | Provide a reusable method that accepts the path of an encrypted .xlsx file and returns its SHA‑256 checksum, handling I/O errors gracefully. | Create a snippet that loads an encrypted workbook using LoadOptions, checks Settings.IsEncrypted, and logs both the encryption status and the computed hash.

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

namespace AsposeCellsHashExample
{
    // Demonstrates how to create a workbook with Aspose.Cells, protect it with a password, save the encrypted file, calculate its SHA‑256 checksum using System.Security.Cryptography, and optionally reload the workbook with LoadOptions to confirm accessibility.
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

            // Save the encrypted workbook to a file
            string encryptedPath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedPath);

            // Compute SHA-256 hash of the encrypted file
            byte[] fileBytes = File.ReadAllBytes(encryptedPath);
            byte[] hashBytes;
            using (SHA256 sha256 = SHA256.Create())
            {
                hashBytes = sha256.ComputeHash(fileBytes);
            }

            // Convert hash to a hexadecimal string for display
            string hashString = BitConverter.ToString(hashBytes).Replace("-", string.Empty);
            Console.WriteLine($"SHA-256 hash of the encrypted workbook: {hashString}");

            // Optional: Load the encrypted workbook to verify it can be opened with the password
            LoadOptions loadOptions = new LoadOptions { Password = "StrongPassword123" };
            Workbook loadedWorkbook = new Workbook(encryptedPath, loadOptions);
            Console.WriteLine($"Workbook loaded successfully. IsEncrypted: {loadedWorkbook.Settings.IsEncrypted}");
        }
    }
}
