// Title: Encrypt an Excel workbook with a password and verify its integrity using SHA-256 checksum in C# (Aspose.Cells)
// Common Searches: C# Aspose.Cells encrypt workbook with password and save as XLSX | How to generate SHA-256 hash of an encrypted Excel file in .NET | Validate integrity of password-protected Excel workbook using checksum in C# | Example code for workbook encryption and checksum verification with Aspose.Cells
// Tags: password encryption of Excel workbook Aspose.Cells | SHA-256 checksum generation C# | integrity verification of encrypted XLSX | Workbook.Save with password Aspose.Cells | file hash comparison after encryption .NET

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

// The sample creates a new workbook, adds data, applies a strong password to encrypt it, saves the file as XLSX, then computes a SHA-256 checksum of the encrypted file. It demonstrates verification by recomputing the hash and comparing the values to ensure the workbook has not been tampered with.
class WorkbookEncryptionWithChecksum
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, Aspose!");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Set a password to encrypt the workbook
        // This will encrypt the file with standard Excel password protection
        workbook.Settings.Password = "StrongPassword123";

        // Define the path for the encrypted file
        string encryptedFilePath = "EncryptedWorkbook.xlsx";

        // Save the workbook (it will be encrypted because a password is set)
        workbook.Save(encryptedFilePath, SaveFormat.Xlsx);

        // Compute a SHA-256 checksum of the encrypted file
        byte[] checksum = ComputeSha256Checksum(encryptedFilePath);

        // Convert checksum to a readable hex string
        string checksumHex = BitConverter.ToString(checksum).Replace("-", string.Empty);
        Console.WriteLine($"SHA-256 checksum of encrypted file: {checksumHex}");

        // Example of verification: load the file and recompute checksum
        byte[] verificationChecksum = ComputeSha256Checksum(encryptedFilePath);
        string verificationHex = BitConverter.ToString(verificationChecksum).Replace("-", string.Empty);

        if (checksumHex.Equals(verificationHex, StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine("Checksum verification passed. File has not been tampered with.");
        }
        else
        {
            Console.WriteLine("Checksum verification failed. File may have been altered.");
        }
    }

    // Helper method to compute SHA-256 checksum of a file
    private static byte[] ComputeSha256Checksum(string filePath)
    {
        using (FileStream stream = File.OpenRead(filePath))
        using (SHA256 sha256 = SHA256.Create())
        {
            return sha256.ComputeHash(stream);
        }
    }
}
