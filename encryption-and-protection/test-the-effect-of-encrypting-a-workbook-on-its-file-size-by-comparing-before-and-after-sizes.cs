// Title: Measure the impact of workbook password encryption on XLSX file size using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells compare XLSX size before and after applying workbook password | how to get file size of encrypted Excel workbook with Aspose.Cells .NET | Aspose.Cells determine size increase caused by workbook encryption in C# | measure password protection overhead on Excel file using Aspose.Cells | retrieve unencrypted and encrypted workbook sizes with Aspose.Cells API
// Tags: Aspose.Cells workbook password encryption size impact | C# XLSX file size after Aspose.Cells encryption | Aspose.Cells Settings.Password file size measurement | compare encrypted vs unencrypted workbook size Aspose.Cells | Excel file size overhead password protection .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, populates it with sample data, saves it as an unencrypted XLSX file, records its byte size, applies a password via wb.Settings.Password, saves the encrypted version, records the new size, and prints the size difference.
class Program
{
    static void Main(string[] args)
    {
        // Define temporary file paths
        string unencryptedPath = Path.Combine(Path.GetTempPath(), "unencrypted.xlsx");
        string encryptedPath = Path.Combine(Path.GetTempPath(), "encrypted.xlsx");

        // -------------------------------------------------
        // Create a new workbook and add sample data
        // -------------------------------------------------
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        // Fill some data to make the file size noticeable
        for (int row = 0; row < 1000; row++)
        {
            for (int col = 0; col < 20; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Save the workbook without encryption
        wb.Save(unencryptedPath, SaveFormat.Xlsx);

        // Get file size of the unencrypted workbook
        long unencryptedSize = new FileInfo(unencryptedPath).Length;

        // -------------------------------------------------
        // Encrypt the same workbook with a password
        // -------------------------------------------------
        // Set password for encryption
        wb.Settings.Password = "SecretPassword123";

        // Save the encrypted workbook
        wb.Save(encryptedPath, SaveFormat.Xlsx);

        // Get file size of the encrypted workbook
        long encryptedSize = new FileInfo(encryptedPath).Length;

        // -------------------------------------------------
        // Output the comparison results
        // -------------------------------------------------
        Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");
        Console.WriteLine($"Encrypted file size:   {encryptedSize} bytes");
        Console.WriteLine($"Size difference:       {encryptedSize - unencryptedSize} bytes");

        // Clean up temporary files (optional)
        // File.Delete(unencryptedPath);
        // File.Delete(encryptedPath);
    }
}
