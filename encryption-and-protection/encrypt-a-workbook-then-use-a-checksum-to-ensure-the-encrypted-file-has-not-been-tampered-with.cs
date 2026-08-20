// Title: Encrypt an Excel workbook with Aspose.Cells and validate it with a SHA‑256 checksum (C#)
// Description: Creates a new Workbook, adds sample data, applies a password and AES‑128 encryption via Aspose.Cells, saves the file, generates a SHA‑256 hash stored in a .txt file, then reloads the protected workbook, recomputes the hash and compares it to detect any tampering.
// Keywords: Aspose.Cells encryption C# | AES-128 Excel protection | SHA256 checksum Excel | password‑protected workbook | integrity verification Aspose.Cells | C# Excel security | EncryptionChecksumDemo GitHub | Excel file tamper detection | secure Excel archive | US developers | EU data protection
// Common Searches: How to encrypt an Excel file with Aspose.Cells and verify its integrity | C# compute SHA‑256 hash for a password‑protected .xlsx | Aspose.Cells set AES 128 encryption options | Load encrypted workbook with password using Aspose.Cells | Validate Excel file checksum after encryption | GitHub Aspose.Cells EncryptionChecksumDemo example
// Developer Intent: Secure an Excel workbook with a password and AES‑128 encryption, then ensure the file remains unchanged by comparing a stored SHA‑256 checksum.
// Use Cases: Distribute confidential financial reports that are encrypted and accompanied by a checksum for client‑side integrity checks. | Automate batch processing of sensitive spreadsheets, rejecting any file whose checksum does not match the original. | Implement a long‑term archival system where each encrypted workbook is paired with a SHA‑256 hash to guarantee data integrity over time.
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, encrypts it with a custom password using AES‑128, and saves a SHA‑256 hash to a text file. | Generate a method that loads a password‑protected workbook with Aspose.Cells and returns true only if the file’s SHA‑256 checksum matches a stored value. | Provide error‑handling patterns for checksum mismatches when opening encrypted Excel files, including logging and throwing a custom exception.

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

namespace AsposeCellsEncryptionChecksumDemo
{
    // Creates a new Workbook, adds sample data, applies a password and AES‑128 encryption via Aspose.Cells, saves the file, generates a SHA‑256 hash stored in a .txt file, then reloads the protected workbook, recomputes the hash and compares it to detect any tampering.
    class Program
    {
        static void Main()
        {
            // Path for the encrypted workbook and checksum file
            string workbookPath = "encrypted.xlsx";
            string checksumPath = "encrypted.sha256";

            // ------------------- Create and encrypt workbook -------------------
            // Create a new workbook (lifecycle rule: create)
            Workbook wb = new Workbook();

            // Add sample data
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data");

            // Set password to encrypt the workbook
            wb.Settings.Password = "StrongPassword123";

            // Optional: set stronger encryption options (e.g., AES 128)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook (lifecycle rule: save)
            wb.Save(workbookPath, SaveFormat.Xlsx);

            // ------------------- Compute and store checksum -------------------
            // Read the saved file bytes
            byte[] fileBytes = File.ReadAllBytes(workbookPath);

            // Compute SHA256 checksum
            byte[] hashBytes;
            using (SHA256 sha256 = SHA256.Create())
            {
                hashBytes = sha256.ComputeHash(fileBytes);
            }

            // Convert checksum to hex string for storage
            string checksumHex = BitConverter.ToString(hashBytes).Replace("-", string.Empty);
            File.WriteAllText(checksumPath, checksumHex);

            Console.WriteLine($"Checksum saved: {checksumHex}");

            // ------------------- Verify checksum after loading -------------------
            // Load the encrypted workbook with password (lifecycle rule: load)
            LoadOptions loadOptions = new LoadOptions { Password = "StrongPassword123" };
            Workbook loadedWb = new Workbook(workbookPath, loadOptions);

            // Re-compute checksum of the file on disk (could also compute from stream)
            byte[] loadedFileBytes = File.ReadAllBytes(workbookPath);
            byte[] loadedHash;
            using (SHA256 sha256 = SHA256.Create())
            {
                loadedHash = sha256.ComputeHash(loadedFileBytes);
            }
            string loadedChecksumHex = BitConverter.ToString(loadedHash).Replace("-", string.Empty);

            // Read the original checksum
            string originalChecksumHex = File.ReadAllText(checksumPath).Trim();

            // Compare checksums
            bool isTampered = !string.Equals(originalChecksumHex, loadedChecksumHex, StringComparison.OrdinalIgnoreCase);
            Console.WriteLine(isTampered
                ? "The encrypted workbook has been tampered with."
                : "Checksum verification passed. The encrypted workbook is intact.");
        }
    }
}
