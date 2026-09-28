// Title: Encrypt an Excel workbook with a password using Aspose.Cells for .NET and generate its SHA‑256 checksum
// AI Prompts: Write C# code that applies an open password to an Aspose.Cells Workbook, saves it as an encrypted XLSX file, and returns the SHA‑256 hash of the saved bytes. | Create a snippet that streams a password‑protected workbook to a MemoryStream, computes the SHA‑256 digest, and outputs the hash in hexadecimal format.
// Common Searches: how to set a password on an Excel file with Aspose.Cells and calculate its SHA256 hash in C# | C# Aspose.Cells encrypt workbook and verify integrity using SHA‑256 checksum | compute SHA‑256 hash of a password‑protected XLSX saved to a MemoryStream using .NET | example of protecting an Excel workbook with Aspose.Cells and generating a hash for the encrypted file
// Tags: Aspose.Cells set workbook password | encrypt workbook to XLSX with Aspose.Cells | SHA-256 hash of encrypted Excel file in .NET | memory stream workbook saving Aspose.Cells | integrity verification of protected XLSX

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

// The program creates a new workbook, adds sample data, applies an open password via Aspose.Cells, saves the encrypted workbook to a MemoryStream in XLSX format, computes the SHA‑256 hash of the encrypted bytes, converts the hash to a hex string, and writes the hash to the console.
class WorkbookEncryptionAndHash
{
    static void Main()
    {
        // Create a new workbook (using the create rule)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Encrypt the workbook with a password (open password)
        workbook.Settings.Password = "StrongPassword123";

        // Save the encrypted workbook to a memory stream (using the save rule)
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsx);
            ms.Position = 0; // Reset stream position for reading

            // Compute SHA‑256 hash of the encrypted workbook bytes
            byte[] hashBytes;
            using (SHA256 sha256 = SHA256.Create())
            {
                hashBytes = sha256.ComputeHash(ms);
            }

            // Convert hash to a hexadecimal string for display
            StringBuilder sb = new StringBuilder();
            foreach (byte b in hashBytes)
                sb.Append(b.ToString("x2"));
            string hashHex = sb.ToString();

            // Output the hash
            Console.WriteLine("SHA‑256 Hash of Encrypted Workbook:");
            Console.WriteLine(hashHex);
        }
    }
}
