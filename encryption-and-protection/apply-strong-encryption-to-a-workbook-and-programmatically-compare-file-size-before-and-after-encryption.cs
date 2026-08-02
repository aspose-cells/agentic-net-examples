// Title: Encrypt an Excel Workbook with Strong 128‑Bit Encryption Using Aspose.Cells (C#) and Compare File Sizes
// Description: Demonstrates how to create a workbook with Aspose.Cells, save it unencrypted, apply a password and 128‑bit strong cryptographic provider, save the encrypted file, and programmatically measure the size difference between the two files.
// Keywords: Aspose.Cells encrypt workbook C# | 128‑bit strong encryption Excel | compare file size before after encryption | set workbook password Aspose.Cells | EncryptionType.StrongCryptographicProvider | measure XLSX size change | C# Excel security example
// Common Searches: How to apply strong 128‑bit encryption to an Excel file with Aspose.Cells | C# code to get size of encrypted vs unencrypted workbook | Aspose.Cells set password and encryption options | Compare XLSX file size before and after encryption | Programmatic Excel file protection using Aspose.Cells .NET
// Developer Intent: Apply strong 128‑bit encryption to a workbook with Aspose.Cells and obtain the before/after file sizes.
// Use Cases: Protect confidential spreadsheets before distribution while quantifying the storage overhead. | Automate compliance checks that require encrypted reports and size tracking. | Integrate encryption and size logging into batch generation of multiple Excel files.
// AI Prompts: Generate C# code that encrypts an existing workbook with 256‑bit AES using Aspose.Cells and reports the size difference. | Explain how different Aspose.Cells encryption settings impact XLSX file size and suggest ways to reduce the increase. | Refactor the sample to process a list of workbooks, encrypt each with a password, and write before/after sizes to a CSV file.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook with Aspose.Cells, save it unencrypted, apply a password and 128‑bit strong cryptographic provider, save the encrypted file, and programmatically measure the size difference between the two files.
class EncryptionComparison
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Information");

        // Save the workbook without encryption
        string unencryptedFile = "Unencrypted.xlsx";
        workbook.Save(unencryptedFile, SaveFormat.Xlsx);

        // Get file size before encryption
        long sizeBefore = new FileInfo(unencryptedFile).Length;

        // Apply strong encryption settings
        workbook.Settings.Password = "StrongPassword123";
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        string encryptedFile = "Encrypted.xlsx";
        workbook.Save(encryptedFile, SaveFormat.Xlsx);

        // Get file size after encryption
        long sizeAfter = new FileInfo(encryptedFile).Length;

        // Output the size comparison
        Console.WriteLine($"Size before encryption: {sizeBefore} bytes");
        Console.WriteLine($"Size after encryption: {sizeAfter} bytes");
        Console.WriteLine($"Size increase: {sizeAfter - sizeBefore} bytes");
    }
}
