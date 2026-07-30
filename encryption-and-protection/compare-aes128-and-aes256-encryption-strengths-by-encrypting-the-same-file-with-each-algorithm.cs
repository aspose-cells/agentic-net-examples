// Title: Encrypt an Excel workbook with AES‑128 vs AES‑256 using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, applies a common password, encrypts it twice—once with AES‑128 and once with AES‑256 via Aspose.Cells' SetEncryptionOptions—saves both files, reloads them to verify encryption, and prints each file's encrypted status and size for a direct strength and overhead comparison.
// Keywords: Aspose.Cells | C# AES encryption | Excel AES-128 | Excel AES-256 | SetEncryptionOptions | StrongCryptographicProvider | password protected workbook | encryption key length | file size comparison | encryption strength example
// Common Searches: How to encrypt an Excel file with AES‑128 using Aspose.Cells C# | How to set AES‑256 encryption for a workbook in .NET | AES‑128 vs AES‑256 file size comparison in Excel | Load password protected workbook with Aspose.Cells C# | Retrieve encryption type from Aspose.Cells workbook
// Developer Intent: Create a workbook, encrypt it with AES‑128 and AES‑256 using Aspose.Cells, then load each file with the password and compare encrypted status and file sizes.
// Use Cases: Validate compliance with policies that require AES‑256 protection. | Show that a single password works for both AES‑128 and AES‑256 encrypted workbooks. | Measure storage overhead introduced by different AES key lengths. | Provide ready‑to‑run sample code for .NET developers adding workbook encryption.
// AI Prompts: Generate C# code that creates an Excel workbook, encrypts it with AES‑128 using Aspose.Cells, saves it, then encrypts the same workbook with AES‑256 and saves the second file. | Write a C# snippet that loads an AES‑encrypted Excel file with Aspose.Cells, checks Settings.IsEncrypted, and outputs the encryption algorithm and key length. | Provide a PowerShell script that computes SHA‑256 hashes of two encrypted Excel files produced by Aspose.Cells and compares them. | Explain the security and performance differences between AES‑128 and AES‑256 encryption in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionComparison
{
    // This example creates a workbook, applies a common password, encrypts it twice—once with AES‑128 and once with AES‑256 via Aspose.Cells' SetEncryptionOptions—saves both files, reloads them to verify encryption, and prints each file's encrypted status and size for a direct strength and overhead comparison.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("AES Encryption Strength Comparison");

            // Common password for both encrypted files
            string password = "Secret123";

            // Set password for the workbook
            workbook.Settings.Password = password;

            // ---------- AES‑128 Encryption ----------
            // Apply AES‑128 (key length = 128 bits) encryption options
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);
            // Save the workbook encrypted with AES‑128
            string file128 = "Encrypted128.xlsx";
            workbook.Save(file128);

            // ---------- AES‑256 Encryption ----------
            // Apply AES‑256 (key length = 256 bits) encryption options
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);
            // Save the workbook encrypted with AES‑256
            string file256 = "Encrypted256.xlsx";
            workbook.Save(file256);

            // Verify that both files are encrypted and can be opened with the password
            LoadOptions loadOptions = new LoadOptions { Password = password };

            Workbook loaded128 = new Workbook(file128, loadOptions);
            Console.WriteLine($"Loaded AES‑128 file. IsEncrypted: {loaded128.Settings.IsEncrypted}");

            Workbook loaded256 = new Workbook(file256, loadOptions);
            Console.WriteLine($"Loaded AES‑256 file. IsEncrypted: {loaded256.Settings.IsEncrypted}");

            // Compare file sizes (larger size may indicate stronger encryption overhead)
            FileInfo info128 = new FileInfo(file128);
            FileInfo info256 = new FileInfo(file256);
            Console.WriteLine($"AES‑128 file size: {info128.Length} bytes");
            Console.WriteLine($"AES‑256 file size: {info256.Length} bytes");
        }
    }
}
