// Title: Password‑protect an Excel workbook and generate a SHA‑256 checksum with Aspose.Cells for .NET (C#)
// Common Searches: how to set a password on an Excel file using Aspose.Cells in C# | C# compute SHA256 hash of a password‑protected XLSX created with Aspose.Cells | save encrypted workbook to a memory stream and get its checksum with Aspose.Cells | verify integrity of a protected Excel workbook using SHA‑256 in .NET | Aspose.Cells encrypt workbook and generate hash for integrity check
// Tags: Aspose.Cells workbook password encryption C# | SHA256 checksum for encrypted XLSX file | memory stream saving of protected Excel workbook Aspose.Cells | integrity verification of password‑protected Excel file .NET | hash generation for encrypted workbook using Aspose.Cells

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

// The example creates a new workbook, adds data, applies a password to encrypt it, saves the encrypted file to a memory stream in XLSX format, computes a SHA‑256 hash of the encrypted bytes for integrity verification, outputs the hash, and optionally writes the encrypted workbook to disk.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Encrypt the workbook with a password
        workbook.Settings.Password = "MySecretPassword";

        // Save the encrypted workbook to a memory stream
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsx);
            byte[] encryptedBytes = ms.ToArray();

            // Compute SHA‑256 hash of the encrypted workbook
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(encryptedBytes);
                string hashString = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                Console.WriteLine("SHA‑256 hash: " + hashString);
            }

            // Optionally write the encrypted file to disk
            File.WriteAllBytes("EncryptedWorkbook.xlsx", encryptedBytes);
        }
    }
}
