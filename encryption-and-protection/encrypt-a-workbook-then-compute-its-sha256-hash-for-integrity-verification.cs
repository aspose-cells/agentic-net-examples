// Title: Encrypt an Excel workbook with a password and generate a SHA‑256 hash using Aspose.Cells for .NET
// Common Searches: C# how to password‑protect an XLSX file with Aspose.Cells and obtain its SHA‑256 checksum | Aspose.Cells encrypt workbook to a memory stream then compute SHA‑256 hash in .NET | verify integrity of a password‑protected Excel file using SHA‑256 in C# | save encrypted workbook as byte array and calculate hash with Aspose.Cells | generate SHA‑256 hash of an encrypted Excel workbook programmatically
// Tags: password protect workbook Aspose.Cells C# | encrypt Excel file to memory stream Aspose.Cells | compute SHA-256 hash of XLSX byte array | integrity verification of encrypted Excel workbook | Aspose.Cells workbook encryption and hashing

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

// The example creates a new Workbook, writes data to cell A1, applies a password to encrypt the file, saves the encrypted workbook to a MemoryStream in XLSX format, computes a SHA‑256 hash of the encrypted bytes, writes the encrypted file to disk, and outputs the hash as a lowercase hexadecimal string.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Optional: add some data
        workbook.Worksheets[0].Cells["A1"].PutValue("Hello Aspose!");

        // Encrypt the workbook with a password
        workbook.Settings.Password = "SecretPassword";

        // Save the encrypted workbook to a memory stream
        using (var ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsx);

            // Compute SHA‑256 hash of the encrypted file bytes
            ms.Position = 0;
            byte[] hashBytes;
            using (var sha256 = SHA256.Create())
            {
                hashBytes = sha256.ComputeHash(ms);
            }

            // Convert hash to a hex string for display
            string hashHex = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();

            // Optionally write the encrypted file to disk
            File.WriteAllBytes("EncryptedWorkbook.xlsx", ms.ToArray());

            Console.WriteLine($"SHA‑256 hash: {hashHex}");
        }
    }
}
