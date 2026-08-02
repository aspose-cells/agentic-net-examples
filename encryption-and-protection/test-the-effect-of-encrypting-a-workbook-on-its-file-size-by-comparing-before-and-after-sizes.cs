// Title: C# – Measure Aspose.Cells Workbook Encryption Overhead on File Size
// Description: Creates a 5,000‑row × 20‑column workbook, saves it unencrypted, applies password protection with StrongCryptographicProvider (128‑bit) encryption, saves the encrypted file, and reports the size difference while verifying the IsEncrypted flag.
// Keywords: Aspose.Cells | C# workbook encryption | Excel file size comparison | SetEncryptionOptions | StrongCryptographicProvider | password protection | IsEncrypted property | encryption overhead
// Common Searches: Aspose.Cells C# encrypt workbook size | how much does Excel encryption increase file size | measure file size difference after password protection Aspose.Cells | SetEncryptionOptions example C# | IsEncrypted flag Aspose.Cells
// Developer Intent: Determine the byte‑level impact of applying password‑based encryption to an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate a large workbook, save it plain, then encrypt it and compare file sizes to quantify overhead. | Load an encrypted workbook with LoadOptions and confirm the IsEncrypted property is true. | Automate size‑impact testing for different encryption algorithms or key lengths in a CI pipeline.
// AI Prompts: Provide a C# script that iterates over all Aspose.Cells encryption types (Standard, StrongCryptographicProvider) and logs file‑size differences for each. | Explain how to extend the example to test both 128‑bit and 256‑bit keys and output a summary table of size changes. | Suggest best practices for measuring encryption overhead in automated tests using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptionSizeTest
{
    // Creates a 5,000‑row × 20‑column workbook, saves it unencrypted, applies password protection with StrongCryptographicProvider (128‑bit) encryption, saves the encrypted file, and reports the size difference while verifying the IsEncrypted flag.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];

            // Populate the worksheet with a reasonable amount of data
            for (int row = 0; row < 5000; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Save the workbook without encryption
            string unencryptedPath = "unencrypted.xlsx";
            wb.Save(unencryptedPath);
            long unencryptedSize = new FileInfo(unencryptedPath).Length;

            // Apply password protection (encryption)
            wb.Settings.Password = "SecretPassword";
            // Optional: specify encryption type and key length
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedPath = "encrypted.xlsx";
            wb.Save(encryptedPath);
            long encryptedSize = new FileInfo(encryptedPath).Length;

            // Output the file sizes
            Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");
            Console.WriteLine($"Encrypted file size:   {encryptedSize} bytes");
            Console.WriteLine($"Size increase:        {encryptedSize - unencryptedSize} bytes");

            // Verify that the encrypted workbook reports IsEncrypted = true
            LoadOptions loadOptions = new LoadOptions { Password = "SecretPassword" };
            Workbook loadedEncrypted = new Workbook(encryptedPath, loadOptions);
            Console.WriteLine($"Loaded workbook IsEncrypted: {loadedEncrypted.Settings.IsEncrypted}");

            // Clean up
            wb.Dispose();
            loadedEncrypted.Dispose();
        }
    }
}
