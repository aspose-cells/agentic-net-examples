// Title: Encrypt an Excel workbook with Aspose.Cells (C#) and verify integrity using SHA‑256 checksum
// Description: This example creates a workbook, writes confidential data, applies a password with Aspose.Cells encryption, saves the file, generates a SHA‑256 hash stored in a separate text file, then reloads the workbook using the password, recomputes the hash and compares it to confirm the file has not been tampered with.
// Keywords: Aspose.Cells encrypt workbook C# | Excel password protection .NET | SHA256 file checksum | verify Excel file integrity | SetEncryptionOptions Aspose | secure Excel export | C# cryptographic hash
// Common Searches: How to password‑protect an Excel file with Aspose.Cells | Generate SHA256 checksum for an encrypted .xlsx in C# | Validate integrity of a protected workbook using Aspose | Aspose.Cells encryption and checksum example | C# code to encrypt Excel and verify hash
// Developer Intent: The developer needs to protect an Excel workbook with a password using Aspose.Cells and ensure the saved file remains unchanged by comparing a stored SHA‑256 checksum.
// Use Cases: Securely distribute financial reports: encrypt the workbook, attach a checksum file, and let recipients verify integrity before opening. | Automate nightly generation of confidential spreadsheets, store a hash for audit trails, and reject any file that fails checksum validation. | Compliance‑driven archiving: encrypt sensitive data, keep a hash for legal proof of unchanged content, and enable quick integrity checks.
// AI Prompts: Write C# code that encrypts an Aspose.Cells workbook with a custom password and saves a SHA256 checksum to a .txt file. | Show how to load a password‑protected workbook with Aspose.Cells and validate its integrity by comparing the stored SHA256 hash. | Explain the effect of SetEncryptionOptions in Aspose.Cells for .xlsx files and how to choose a strong encryption algorithm.

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

namespace AsposeCellsEncryptionChecksumDemo
{
    // This example creates a workbook, writes confidential data, applies a password with Aspose.Cells encryption, saves the file, generates a SHA‑256 hash stored in a separate text file, then reloads the workbook using the password, recomputes the hash and compares it to confirm the file has not been tampered with.
    class Program
    {
        // Path for the encrypted workbook
        private const string EncryptedFilePath = "encrypted.xlsx";
        // Path for storing the checksum
        private const string ChecksumFilePath = "encrypted_checksum.txt";
        // Password used for encryption
        private const string WorkbookPassword = "mySecretPassword";

        static void Main()
        {
            // ---------- Create and encrypt the workbook ----------
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data");
            // Set password to encrypt the file
            wb.Settings.Password = WorkbookPassword;
            // Optional: set stronger encryption options (ignored for .xlsx but kept for completeness)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);
            // Save the encrypted workbook
            wb.Save(EncryptedFilePath);
            // Compute and store checksum of the encrypted file
            string checksum = ComputeFileChecksum(EncryptedFilePath);
            File.WriteAllText(ChecksumFilePath, checksum);
            Console.WriteLine($"Workbook encrypted and saved to '{EncryptedFilePath}'.");
            Console.WriteLine($"Checksum stored in '{ChecksumFilePath}': {checksum}");

            // ---------- Verify integrity using checksum ----------
            // Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions { Password = WorkbookPassword };
            Workbook loadedWb = new Workbook(EncryptedFilePath, loadOptions);
            // Re‑compute checksum of the file on disk
            string currentChecksum = ComputeFileChecksum(EncryptedFilePath);
            // Read the original checksum
            string originalChecksum = File.ReadAllText(ChecksumFilePath);
            // Compare
            if (string.Equals(currentChecksum, originalChecksum, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Checksum verification passed. File has not been tampered with.");
            }
            else
            {
                Console.WriteLine("Checksum verification failed! The file may have been altered.");
            }

            // Optional: demonstrate that data can be read after decryption
            Console.WriteLine("Decrypted cell A1 value: " + loadedWb.Worksheets[0].Cells["A1"].Value);
        }

        // Helper method to compute SHA256 checksum of a file and return as hex string
        private static string ComputeFileChecksum(string filePath)
        {
            using (FileStream stream = File.OpenRead(filePath))
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", string.Empty);
            }
        }
    }
}
