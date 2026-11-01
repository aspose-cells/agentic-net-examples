// Title: Encrypt an Excel workbook with a password and validate its integrity using a SHA‑256 checksum in C# (Aspose.Cells)
// Common Searches: C# Aspose.Cells encrypt workbook with password and save to memory stream | How to generate SHA256 hash of an encrypted XLSX file using .NET | Verify integrity of password‑protected Excel file after saving with Aspose.Cells | Compute checksum of workbook after applying workbook.Settings.Password in C# | Aspose.Cells example for workbook encryption and checksum verification
// Tags: Aspose.Cells workbook password encryption C# | SHA256 checksum generation for XLSX files | encrypt workbook to memory stream Aspose.Cells | integrity verification of encrypted Excel using .NET | password protected Excel file checksum Aspose

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

// The sample creates a new workbook, sets a strong password via workbook.Settings.Password, saves the encrypted workbook to a MemoryStream as XLSX, computes a SHA‑256 hash of the encrypted bytes, writes the file to disk, then recomputes the hash to confirm the file has not been altered.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample");
        sheet.Cells["B1"].PutValue("Data");

        // Encrypt the workbook with a password
        workbook.Settings.Password = "StrongPassword123";

        // Save the encrypted workbook to a memory stream
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsx);
            byte[] encryptedBytes = ms.ToArray();

            // Compute SHA256 checksum of the encrypted file
            string checksum = ComputeChecksum(encryptedBytes);
            Console.WriteLine($"Checksum: {checksum}");

            // Write the encrypted workbook to disk
            File.WriteAllBytes("EncryptedWorkbook.xlsx", encryptedBytes);

            // Example verification: recompute checksum and compare
            byte[] fileBytes = File.ReadAllBytes("EncryptedWorkbook.xlsx");
            string verifyChecksum = ComputeChecksum(fileBytes);
            Console.WriteLine($"Checksum verification passed: {checksum == verifyChecksum}");
        }
    }

    // Helper method to compute SHA256 checksum
    static string ComputeChecksum(byte[] data)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hash = sha256.ComputeHash(data);
            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }
    }
}
