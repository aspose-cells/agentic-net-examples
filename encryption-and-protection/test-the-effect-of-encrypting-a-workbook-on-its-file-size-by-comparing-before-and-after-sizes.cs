// Title: Determine how workbook password encryption affects XLSX file size with Aspose.Cells for .NET
// AI Prompts: Write a C# program using Aspose.Cells that creates a workbook, saves it as an unencrypted XLSX, then sets Workbook.Settings.Password, saves an encrypted copy, and outputs both file sizes. | Extend the program to compute and display the percentage growth of the XLSX file after password protection. | Refactor the example into a reusable method that accepts a password string and returns the size difference between the encrypted and original workbook.
// Common Searches: how much does Aspose.Cells password encryption increase XLSX size in C# | C# code to get original and encrypted Excel file sizes with Aspose.Cells | measure percentage change of XLSX file after applying workbook password using Aspose.Cells | determine file size impact of Settings.Password on Excel workbook in .NET
// Tags: Aspose.Cells workbook encryption size measurement | C# XLSX file size before and after encryption | measure encrypted Excel workbook size .NET | compare protected vs unprotected workbook file size

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating a workbook with sample data, saving it as an unencrypted XLSX, applying a password via Workbook.Settings.Password, saving the encrypted file, and printing the original size, encrypted size, and size increase in bytes.
class WorkbookEncryptionSizeTest
{
    static void Main()
    {
        // Path for the original (unencrypted) workbook
        string originalPath = "original.xlsx";
        // Path for the encrypted workbook
        string encryptedPath = "encrypted.xlsx";

        // Create a new workbook and add some sample data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        // Fill the sheet with data to make the file size noticeable
        for (int row = 0; row < 1000; row++)
        {
            for (int col = 0; col < 20; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Save the workbook without encryption
        wb.Save(originalPath, SaveFormat.Xlsx);

        // Get file size of the unencrypted workbook
        long originalSize = new FileInfo(originalPath).Length;

        // Apply password protection (encryption)
        wb.Settings.Password = "SecretPassword";

        // Save the encrypted workbook
        wb.Save(encryptedPath, SaveFormat.Xlsx);

        // Get file size of the encrypted workbook
        long encryptedSize = new FileInfo(encryptedPath).Length;

        // Output the sizes for comparison
        Console.WriteLine($"Original (unencrypted) size: {originalSize} bytes");
        Console.WriteLine($"Encrypted size: {encryptedSize} bytes");
        Console.WriteLine($"Size increase: {encryptedSize - originalSize} bytes");
    }
}
