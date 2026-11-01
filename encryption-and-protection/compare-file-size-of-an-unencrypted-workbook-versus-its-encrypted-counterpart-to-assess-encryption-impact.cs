// Title: Compare unencrypted and password‑protected Excel workbook file sizes with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get size of .xlsx before and after applying password protection | how much does password encryption increase Excel file size using Aspose.Cells | measure encryption overhead of an Aspose.Cells workbook in .NET | compare file size of encrypted vs unencrypted workbook with Aspose.Cells C# example | Aspose.Cells calculate size difference between protected and unprotected Excel files
// Tags: Aspose.Cells workbook encryption size analysis | C# .xlsx password protection file size impact | compare encrypted and unencrypted workbook sizes Aspose.Cells | measure Aspose.Cells encryption overhead .NET | file size difference after Aspose.Cells password protect

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a workbook, saves it as an unencrypted .xlsx file, records its size, then applies a password via Workbook.Settings.Password, saves the encrypted version, records its size, and prints both sizes along with the size increase to illustrate the encryption overhead.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Text");
        sheet.Cells["B2"].PutValue(12345);

        // Save the workbook without encryption
        string unencryptedPath = "Unencrypted.xlsx";
        workbook.Save(unencryptedPath, SaveFormat.Xlsx);
        long unencryptedSize = new FileInfo(unencryptedPath).Length;

        // Apply password protection (encryption)
        workbook.Settings.Password = "MySecretPassword";

        // Save the encrypted workbook
        string encryptedPath = "Encrypted.xlsx";
        workbook.Save(encryptedPath, SaveFormat.Xlsx);
        long encryptedSize = new FileInfo(encryptedPath).Length;

        // Compare and display file sizes
        Console.WriteLine($"Unencrypted size: {unencryptedSize} bytes");
        Console.WriteLine($"Encrypted size:   {encryptedSize} bytes");
        Console.WriteLine($"Size increase:    {encryptedSize - unencryptedSize} bytes");
    }
}
