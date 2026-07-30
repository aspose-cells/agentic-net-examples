// Title: Aspose.Cells C# – Compare Unencrypted and Encrypted Workbook File Sizes
// Description: Creates a 1000‑row × 10‑column workbook, saves it as an unencrypted XLSX, records the size, applies a password with 128‑bit strong encryption, saves the encrypted file, records the new size, and loads the encrypted workbook to confirm the IsEncrypted flag.
// Keywords: Aspose.Cells | C# | .NET | encrypt workbook | file size comparison | password protection | 128‑bit encryption | EncryptionType.StrongCryptographicProvider | measure Excel size overhead | load encrypted workbook
// Common Searches: Aspose.Cells file size after password protection | C# compare encrypted vs unencrypted Excel size | how much overhead does Aspose.Cells encryption add | verify encrypted workbook with Aspose.Cells LoadOptions | measure Excel workbook size difference in .NET
// Developer Intent: Find out how Aspose.Cells encryption influences the physical size of an Excel file.
// Use Cases: Generate a large workbook, save it plain, then encrypt with a password and 128‑bit encryption, and log both file sizes for storage impact analysis. | Load an encrypted workbook using LoadOptions to ensure the IsEncrypted property returns true. | Automate reporting of size overhead for compliance or capacity‑planning scripts that produce password‑protected Excel files.
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, saves an unencrypted copy, encrypts it with a 128‑bit password, saves the encrypted copy, and prints the size difference. | Explain why Aspose.Cells encryption may increase file size and suggest techniques to reduce the overhead. | Generate a C# unit test that asserts the encrypted XLSX size is greater than or equal to the unencrypted size when using Aspose.Cells encryption options.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionSizeDemo
{
    // Creates a 1000‑row × 10‑column workbook, saves it as an unencrypted XLSX, records the size, applies a password with 128‑bit strong encryption, saves the encrypted file, records the new size, and loads the encrypted workbook to confirm the IsEncrypted flag.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 1000; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Save the workbook without encryption
            string unencryptedPath = "unencrypted.xlsx";
            workbook.Save(unencryptedPath);
            long unencryptedSize = new FileInfo(unencryptedPath).Length;

            // Apply password protection (encryption)
            workbook.Settings.Password = "SecretPassword";
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedPath = "encrypted.xlsx";
            workbook.Save(encryptedPath);
            long encryptedSize = new FileInfo(encryptedPath).Length;

            // Output file sizes
            Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");
            Console.WriteLine($"Encrypted file size:   {encryptedSize} bytes");

            // Verify that the encrypted file is indeed encrypted
            LoadOptions loadOptions = new LoadOptions { Password = "SecretPassword" };
            Workbook loadedEncrypted = new Workbook(encryptedPath, loadOptions);
            Console.WriteLine($"Loaded workbook IsEncrypted: {loadedEncrypted.Settings.IsEncrypted}");

            // Clean up
            workbook.Dispose();
            loadedEncrypted.Dispose();
        }
    }
}
