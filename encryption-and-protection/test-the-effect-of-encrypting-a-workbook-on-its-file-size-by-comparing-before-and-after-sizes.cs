// Title: Determine how workbook encryption affects XLSX file size with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells compare XLSX size before and after applying password protection | measure file size increase caused by workbook encryption using Aspose.Cells .NET | how to get encrypted workbook file length in Aspose.Cells C# | Aspose.Cells .NET check encryption overhead on generated Excel file | sample code to save encrypted Excel and read its size with Aspose.Cells
// Tags: Aspose.Cells workbook password protection size impact | XLSX encryption overhead Aspose.Cells .NET | measure encrypted Excel file size C# | SaveFormat.Xlsx with password Aspose.Cells | file size comparison unencrypted encrypted workbook

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook with sample data, saves it as an unencrypted XLSX, records its size, applies a password via Workbook.Settings.Password to encrypt it, saves the encrypted file, records the new size, and writes both sizes to the console.
class WorkbookEncryptionSizeTest
{
    static void Main()
    {
        // Create a new workbook and add some data to increase its size
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        // Fill the sheet with sample data (e.g., 1000 rows x 10 columns)
        for (int row = 0; row < 1000; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Define file paths for the unencrypted and encrypted workbooks
        string unencryptedPath = Path.Combine(Path.GetTempPath(), "UnencryptedWorkbook.xlsx");
        string encryptedPath = Path.Combine(Path.GetTempPath(), "EncryptedWorkbook.xlsx");

        // Save the workbook without encryption
        wb.Save(unencryptedPath, SaveFormat.Xlsx);

        // Get file size before encryption
        long sizeBefore = new FileInfo(unencryptedPath).Length;

        // Apply password protection (encryption)
        wb.Settings.Password = "SecretPassword";

        // Save the workbook with encryption
        wb.Save(encryptedPath, SaveFormat.Xlsx);

        // Get file size after encryption
        long sizeAfter = new FileInfo(encryptedPath).Length;

        // Output the results
        Console.WriteLine($"File size before encryption:  {sizeBefore} bytes");
        Console.WriteLine($"File size after encryption:   {sizeAfter} bytes");

        // Clean up temporary files (optional)
        // File.Delete(unencryptedPath);
        // File.Delete(encryptedPath);
    }
}
