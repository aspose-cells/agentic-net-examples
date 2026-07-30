// Title: Compute SHA‑256 Hash of a Password‑Protected Aspose.Cells Workbook (C#)
// Description: Creates a workbook with Aspose.Cells, applies a password, saves it, then streams the encrypted file and uses System.Security.Cryptography.SHA256 to generate a hexadecimal hash for integrity verification.
// Keywords: Aspose.Cells SHA256 hash | C# encrypted Excel checksum | password protected workbook hash | verify Excel file integrity .NET | compute file hash Aspose.Cells
// Common Searches: C# compute SHA-256 of encrypted Excel file | Aspose.Cells password protection hash | verify integrity of password protected workbook | how to get checksum of encrypted .xlsx using .NET | Aspose.Cells SHA256 example
// Developer Intent: Generate a SHA‑256 checksum for a workbook saved with Aspose.Cells password protection to confirm that the file has not been altered.
// Use Cases: Store the hash alongside the encrypted workbook for later tamper detection. | Log the checksum in audit trails to satisfy compliance requirements. | Compare a downloaded encrypted workbook's hash with an expected value before processing.
// AI Prompts: Provide C# code that saves an Aspose.Cells workbook with a password and then calculates its SHA‑256 hash as a hex string. | Show how to verify the integrity of a password‑protected Excel file by comparing its SHA‑256 hash to a known value using .NET cryptography. | Explain an efficient streaming approach to compute a SHA‑256 hash for large encrypted workbooks without loading the entire file into memory.

using Aspose.Cells;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

// Creates a workbook with Aspose.Cells, applies a password, saves it, then streams the encrypted file and uses System.Security.Cryptography.SHA256 to generate a hexadecimal hash for integrity verification.
class ComputeWorkbookHash
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");

        // Encrypt the workbook with a password
        workbook.Settings.Password = "mySecretPassword";

        // Save the encrypted workbook to disk
        string filePath = "encryptedWorkbook.xlsx";
        workbook.Save(filePath);

        // Compute SHA‑256 hash of the saved encrypted file
        byte[] hashBytes;
        using (FileStream stream = File.OpenRead(filePath))
        using (SHA256 sha256 = SHA256.Create())
        {
            hashBytes = sha256.ComputeHash(stream);
        }

        // Convert the hash bytes to a hexadecimal string for display
        StringBuilder sb = new StringBuilder();
        foreach (byte b in hashBytes)
        {
            sb.Append(b.ToString("x2"));
        }
        string hashString = sb.ToString();

        Console.WriteLine($"SHA‑256 hash of the encrypted workbook: {hashString}");
    }
}
