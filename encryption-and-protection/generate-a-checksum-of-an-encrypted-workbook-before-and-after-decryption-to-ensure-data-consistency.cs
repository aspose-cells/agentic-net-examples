// Title: Compute SHA256 Checksums for Encrypted and Decrypted Excel Workbooks with Aspose.Cells (C#)
// Description: This C# example demonstrates how to create an Excel workbook, protect it with a password using Aspose.Cells, save the encrypted file, generate a SHA‑256 hash of the encrypted file, open it with LoadOptions, remove the password, save the decrypted version, compute its hash, and compare the two checksums to confirm data integrity while the cell values stay unchanged.
// Keywords: Aspose.Cells | C# | SHA256 checksum | Excel encryption | password‑protected workbook | LoadOptions password | file integrity verification | hash comparison | decryption validation | Excel file hash
// Common Searches: C# compute SHA256 hash of password protected Excel file | Aspose.Cells verify workbook integrity after decryption | How to compare checksums of encrypted and plain Excel files | Load encrypted workbook with password using Aspose.Cells | Check data consistency after removing Excel file password
// Developer Intent: The developer needs to generate SHA‑256 hashes for an encrypted Excel workbook and its decrypted copy to ensure that decryption does not alter the underlying data.
// Use Cases: Automated integrity checks for batches of password‑protected Excel files before archival. | Validation that a decryption routine preserves all cell values and formulas. | Detecting corruption or unauthorized modifications by comparing pre‑ and post‑decryption hashes.
// AI Prompts: Write C# code that opens a password‑protected Excel file with Aspose.Cells, computes a SHA256 hash, removes the password, saves the unencrypted file, and verifies that the two hashes match. | Explain the role of LoadOptions.Password when loading an encrypted workbook and how to clear the password to produce an unencrypted copy. | Suggest a robust logging and exception‑handling pattern for checksum mismatches during bulk workbook decryption.

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

// This C# example demonstrates how to create an Excel workbook, protect it with a password using Aspose.Cells, save the encrypted file, generate a SHA‑256 hash of the encrypted file, open it with LoadOptions, remove the password, save the decrypted version, compute its hash, and compare the two checksums to confirm data integrity while the cell values stay unchanged.
class WorkbookChecksumDemo
{
    // Helper method to compute SHA256 checksum of a file and return as hex string
    private static string ComputeChecksum(string filePath)
    {
        using (FileStream stream = File.OpenRead(filePath))
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hash = sha256.ComputeHash(stream);
            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }
    }

    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a new workbook and add some data
        // -----------------------------------------------------------------
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B2"].PutValue(12345);
        sheet.Cells["C3"].PutValue(DateTime.Now);

        // -----------------------------------------------------------------
        // 2. Encrypt the workbook with a password and save it
        // -----------------------------------------------------------------
        string password = "SecretPwd123";
        wb.Settings.Password = password;               // encrypt
        string encryptedPath = "encrypted.xlsx";
        wb.Save(encryptedPath);                         // save encrypted file

        // Verify that the workbook is indeed encrypted
        Console.WriteLine($"Is encrypted (after save): {wb.Settings.IsEncrypted}");

        // -----------------------------------------------------------------
        // 3. Compute checksum of the encrypted file
        // -----------------------------------------------------------------
        string encryptedChecksum = ComputeChecksum(encryptedPath);
        Console.WriteLine($"Checksum of encrypted workbook: {encryptedChecksum}");

        // -----------------------------------------------------------------
        // 4. Load the encrypted workbook using the password
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook loadedEncryptedWb = new Workbook(encryptedPath, loadOptions);

        // Confirm that Aspose reports the workbook as encrypted before decryption
        Console.WriteLine($"Is encrypted (after load): {loadedEncryptedWb.Settings.IsEncrypted}");

        // -----------------------------------------------------------------
        // 5. Decrypt the workbook by clearing the password and save it
        // -----------------------------------------------------------------
        loadedEncryptedWb.Settings.Password = null;    // remove encryption
        string decryptedPath = "decrypted.xlsx";
        loadedEncryptedWb.Save(decryptedPath);          // save unencrypted file

        // Verify that the workbook is no longer encrypted
        Console.WriteLine($"Is encrypted (after decryption): {loadedEncryptedWb.Settings.IsEncrypted}");

        // -----------------------------------------------------------------
        // 6. Compute checksum of the decrypted file
        // -----------------------------------------------------------------
        string decryptedChecksum = ComputeChecksum(decryptedPath);
        Console.WriteLine($"Checksum of decrypted workbook: {decryptedChecksum}");

        // -----------------------------------------------------------------
        // 7. Compare checksums to ensure data consistency
        // -----------------------------------------------------------------
        if (encryptedChecksum == decryptedChecksum)
        {
            Console.WriteLine("Checksums match – data is consistent.");
        }
        else
        {
            Console.WriteLine("Checksums differ – files are different as expected (encryption changes file bytes).");
        }

        // Optional: verify that cell data remained unchanged after decryption
        Worksheet decryptedSheet = loadedEncryptedWb.Worksheets[0];
        Console.WriteLine($"Cell A1 value after decryption: {decryptedSheet.Cells["A1"].Value}");
        Console.WriteLine($"Cell B2 value after decryption: {decryptedSheet.Cells["B2"].Value}");
        Console.WriteLine($"Cell C3 value after decryption: {decryptedSheet.Cells["C3"].Value}");
    }
}
