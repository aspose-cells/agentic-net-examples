// Title: Compute and compare SHA256 checksums of an encrypted Aspose.Cells workbook before and after decryption in C#
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, saves it with a password, calculates its SHA256 hash, then opens it with LoadOptions, removes the password, re‑saves it, and verifies that the two hashes are identical. | Write a reusable method that accepts a MemoryStream of an Excel file and returns a SHA256 checksum, and demonstrate its use to ensure data consistency of a password‑protected workbook after decryption with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to get hash of an encrypted XLSX file | verify integrity of password protected Excel workbook using Aspose.Cells | C# compute workbook stream hash before and after removing password with Aspose.Cells | load encrypted Excel file with password in Aspose.Cells and compare file hashes | Aspose.Cells hash mismatch after decryption troubleshooting
// Tags: hash calculation for Aspose.Cells workbook | password protected XLSX integrity check using Aspose.Cells | load options with password decryption Aspose.Cells C# | memory stream hash comparison Aspose.Cells | encrypted workbook validation C#

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

// // Demonstrates creating a workbook, protecting it with a password, computing a SHA256 hash of the encrypted stream, decrypting it via LoadOptions, recomputing the hash of the decrypted stream, and comparing the hashes to confirm data integrity after decryption.
class WorkbookChecksumDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B2"].PutValue(12345);
        sheet.Cells["C3"].PutValue(DateTime.Now);

        // Define password for encryption
        string password = "SecretPwd";

        // Save the workbook as an encrypted file into a memory stream
        MemoryStream encryptedStream = new MemoryStream();
        // Apply password protection
        wb.Settings.Password = password;
        wb.Save(encryptedStream, SaveFormat.Xlsx);
        // Reset stream position for reading
        encryptedStream.Position = 0;

        // Compute checksum (SHA256) of the encrypted workbook bytes
        byte[] encryptedBytes = encryptedStream.ToArray();
        string encryptedChecksum = ComputeSha256Hash(encryptedBytes);
        Console.WriteLine($"Encrypted workbook checksum: {encryptedChecksum}");

        // Load the encrypted workbook using the password (decryption)
        LoadOptions loadOpts = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = password
        };
        Workbook decryptedWb = new Workbook(encryptedStream, loadOpts);

        // Save the decrypted workbook (without password) into another memory stream
        MemoryStream decryptedStream = new MemoryStream();
        // Ensure no password is set for the output
        decryptedWb.Settings.Password = null;
        decryptedWb.Save(decryptedStream, SaveFormat.Xlsx);
        decryptedStream.Position = 0;

        // Compute checksum (SHA256) of the decrypted workbook bytes
        byte[] decryptedBytes = decryptedStream.ToArray();
        string decryptedChecksum = ComputeSha256Hash(decryptedBytes);
        Console.WriteLine($"Decrypted workbook checksum: {decryptedChecksum}");

        // Verify data consistency
        if (encryptedChecksum == decryptedChecksum)
        {
            Console.WriteLine("Checksums match: data is consistent after decryption.");
        }
        else
        {
            Console.WriteLine("Checksums differ: data inconsistency detected.");
        }
    }

    // Helper method to compute SHA256 hash and return as hex string
    private static string ComputeSha256Hash(byte[] data)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(data);
            // Convert hash bytes to hex string
            return BitConverter.ToString(hashBytes).Replace("-", string.Empty);
        }
    }
}
