// Title: Encrypt an Excel Workbook with Strong 128‑Bit Cryptography and Compare File Sizes using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, save it unencrypted, apply a password and 128‑bit StrongCryptographicProvider encryption, save the protected file, and programmatically read and display the original and encrypted file sizes to illustrate the size impact.
// Keywords: Aspose.Cells encryption C# | StrongCryptographicProvider | 128‑bit Excel encryption | workbook password protection | file size comparison after encryption | .NET Excel security | measure encrypted file size
// Common Searches: how to encrypt an Excel file with Aspose.Cells .NET | set 128‑bit password protection for workbook using C# | compare Excel file size before and after encryption | Aspose.Cells StrongCryptographicProvider example | measure size increase of encrypted Excel workbook
// Developer Intent: Apply strong 128‑bit encryption to an Excel workbook and programmatically compare its size before and after protection.
// Use Cases: Secure confidential spreadsheets before distribution by adding a password and strong encryption. | Validate compliance requirements by automatically checking the size delta between original and encrypted files. | Generate audit logs that capture both unencrypted and encrypted workbook sizes for reporting.
// AI Prompts: Provide C# code that encrypts an existing Excel file with a 256‑bit key using Aspose.Cells and outputs the size difference. | Show how to handle exceptions when setting encryption options and saving an encrypted workbook in Aspose.Cells. | Create a script that encrypts multiple workbooks, records original and encrypted sizes in a CSV, and logs any failures.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, save it unencrypted, apply a password and 128‑bit StrongCryptographicProvider encryption, save the protected file, and programmatically read and display the original and encrypted file sizes to illustrate the size impact.
class StrongEncryptionDemo
{
    static void Main()
    {
        // Paths for the unencrypted and encrypted workbooks
        string unencryptedPath = "UnencryptedWorkbook.xlsx";
        string encryptedPath = "EncryptedWorkbook.xlsx";

        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for encryption test.");

        // -------------------------------------------------
        // Save the workbook without any protection (unencrypted)
        // -------------------------------------------------
        workbook.Save(unencryptedPath);
        long unencryptedSize = new FileInfo(unencryptedPath).Length;

        // -------------------------------------------------
        // Apply strong encryption settings
        // -------------------------------------------------
        // Set a password that will be required to open the file
        workbook.Settings.Password = "StrongPassword123";

        // Set encryption options: StrongCryptographicProvider with 128‑bit key
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        workbook.Save(encryptedPath);
        long encryptedSize = new FileInfo(encryptedPath).Length;

        // -------------------------------------------------
        // Compare file sizes before and after encryption
        // -------------------------------------------------
        Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");
        Console.WriteLine($"Encrypted file size:   {encryptedSize} bytes");
        Console.WriteLine($"Size increase:         {encryptedSize - unencryptedSize} bytes");

        // Clean up
        workbook.Dispose();
    }
}
