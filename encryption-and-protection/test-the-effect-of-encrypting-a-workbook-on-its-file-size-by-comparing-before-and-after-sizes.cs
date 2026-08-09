// Title: C# – Measure Workbook Size Change After AES‑128 Encryption with Aspose.Cells
// Description: Creates an XLSX file, records its size, applies password‑protected AES‑128 encryption via Aspose.Cells, saves the encrypted version, and outputs both sizes to show the encryption overhead.
// Keywords: Aspose.Cells | C# | .NET | workbook encryption | AES-128 | file size comparison | password protection | Excel size impact | encryption overhead
// Common Searches: Aspose.Cells how much does AES encryption increase Excel file size | C# compare encrypted vs unencrypted workbook size | measure encryption overhead for Excel files using Aspose.Cells | file size difference after applying password protection in .NET
// Developer Intent: Find out how password‑protected AES‑128 encryption influences the size of an Excel workbook generated with Aspose.Cells.
// Use Cases: Benchmark encryption overhead for large reports before deployment. | Validate that encrypted workbooks meet storage‑budget constraints. | Automate size logging for compliance‑driven documents that require password protection.
// AI Prompts: Generate C# code that encrypts a workbook with AES‑256 using Aspose.Cells and prints the percentage size increase. | Explain how to retrieve the encryption algorithm and compute the exact byte difference between encrypted and original files in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Creates an XLSX file, records its size, applies password‑protected AES‑128 encryption via Aspose.Cells, saves the encrypted version, and outputs both sizes to show the encryption overhead.
class WorkbookEncryptionSizeDemo
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for encryption size test.");

        // Save the workbook without encryption
        string unencryptedFile = "Unencrypted.xlsx";
        workbook.Save(unencryptedFile);
        long unencryptedSize = new FileInfo(unencryptedFile).Length;

        // Apply password protection (encryption)
        workbook.Settings.Password = "myPassword";
        // Use strong encryption (AES 128) for the workbook
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        string encryptedFile = "Encrypted.xlsx";
        workbook.Save(encryptedFile);
        long encryptedSize = new FileInfo(encryptedFile).Length;

        // Output the file sizes for comparison
        Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");
        Console.WriteLine($"Encrypted file size: {encryptedSize} bytes");

        // Clean up resources
        workbook.Dispose();
    }
}
