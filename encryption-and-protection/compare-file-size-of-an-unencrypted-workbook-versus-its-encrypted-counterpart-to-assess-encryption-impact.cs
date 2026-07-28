// Title: C# Example: Compare Unencrypted vs Password‑Protected Excel File Size with Aspose.Cells
// Description: Creates a simple workbook, saves it as an unencrypted XLSX, records its size, then applies a password with AES‑128 encryption, saves the encrypted file, and prints both sizes and the size increase. Demonstrates how to measure encryption overhead using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# encryption | Excel file size after password protection | measure encryption overhead .NET | compare encrypted vs unencrypted workbook | AES 128 Aspose.Cells example | file size impact password protection | C# Aspose.Cells sample code | GitHub Aspose.Cells encryption demo
// Common Searches: Aspose.Cells compare encrypted workbook size | C# get size of password‑protected Excel file | how much does encryption increase XLSX size | measure file size difference after applying Aspose.Cells encryption | C# example for Aspose.Cells file size comparison
// Developer Intent: Find out how much an Excel workbook grows when saved with password encryption using Aspose.Cells.
// Use Cases: Benchmark storage impact before enabling workbook protection in reporting pipelines. | Document encryption overhead for compliance or audit trails. | Automate size‑difference logging for large batches of spreadsheets in cloud storage.
// AI Prompts: Generate C# code with Aspose.Cells that calculates the percentage increase in file size after applying AES‑256 encryption. | Show how to process a directory of Excel files, save each both unencrypted and encrypted, and export a CSV with size comparisons. | Explain the effect of different encryption algorithms (AES‑128, AES‑256, RC4) on XLSX file size when using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace EncryptionImpactDemo
{
    // Creates a simple workbook, saves it as an unencrypted XLSX, records its size, then applies a password with AES‑128 encryption, saves the encrypted file, and prints both sizes and the size increase. Demonstrates how to measure encryption overhead using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for encryption impact analysis.");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue(12345);

            // Define file paths for unencrypted and encrypted workbooks
            string unencryptedPath = "Unencrypted.xlsx";
            string encryptedPath = "Encrypted.xlsx";

            // Save the workbook without any protection (unencrypted)
            wb.Save(unencryptedPath, SaveFormat.Xlsx);

            // Get file size of the unencrypted workbook
            long unencryptedSize = new FileInfo(unencryptedPath).Length;

            // Apply password protection (encryption) to the same workbook
            wb.Settings.Password = "StrongPassword123";
            // Optionally set encryption options (e.g., AES 128-bit)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            wb.Save(encryptedPath, SaveFormat.Xlsx);

            // Get file size of the encrypted workbook
            long encryptedSize = new FileInfo(encryptedPath).Length;

            // Output the sizes and the impact of encryption
            Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");
            Console.WriteLine($"Encrypted file size:   {encryptedSize} bytes");
            Console.WriteLine($"Size increase:        {encryptedSize - unencryptedSize} bytes");
        }
    }
}
