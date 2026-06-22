using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

class WorkbookEncryptionWithChecksum
{
    static void Main()
    {
        // Path for the encrypted workbook and checksum file
        string workbookPath = "encrypted.xlsx";
        string checksumPath = "encrypted_checksum.txt";

        // ------------------- Create and encrypt workbook -------------------
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add some data to the first worksheet
        wb.Worksheets[0].Cells["A1"].PutValue("Sensitive data that needs protection.");

        // Set a password to encrypt the workbook
        wb.Settings.Password = "mySecretPassword";

        // Save the encrypted workbook
        wb.Save(workbookPath);

        // ------------------- Compute and store checksum -------------------
        // Read the saved file bytes
        byte[] fileBytes = File.ReadAllBytes(workbookPath);

        // Compute SHA256 hash
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hash = sha256.ComputeHash(fileBytes);
            string hashString = BitConverter.ToString(hash).Replace("-", string.Empty);

            // Store the checksum in a separate text file
            File.WriteAllText(checksumPath, hashString);
        }

        // ------------------- Load workbook and verify checksum -------------------
        // Load options with the password to open the encrypted workbook
        LoadOptions loadOptions = new LoadOptions { Password = "mySecretPassword" };
        Workbook loadedWb = new Workbook(workbookPath, loadOptions);

        // Re‑compute checksum of the loaded file
        byte[] loadedFileBytes = File.ReadAllBytes(workbookPath);
        string recomputedHash;
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hash = sha256.ComputeHash(loadedFileBytes);
            recomputedHash = BitConverter.ToString(hash).Replace("-", string.Empty);
        }

        // Read the original checksum
        string originalHash = File.ReadAllText(checksumPath).Trim();

        // Compare hashes to detect tampering
        if (string.Equals(originalHash, recomputedHash, StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine("Checksum verification passed. The encrypted workbook has not been tampered with.");
        }
        else
        {
            Console.WriteLine("Checksum verification failed! The encrypted workbook may have been altered.");
        }

        // Optional: demonstrate that data can be accessed after successful decryption
        Console.WriteLine("Cell A1 value after decryption: " + loadedWb.Worksheets[0].Cells["A1"].Value);
    }
}