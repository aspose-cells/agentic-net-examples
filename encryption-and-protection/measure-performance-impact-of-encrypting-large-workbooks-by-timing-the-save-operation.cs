// Title: How to measure the performance impact of password‑protecting a large Excel workbook when saving with Aspose.Cells for .NET
// Common Searches: C# code to benchmark save time of a password‑protected Excel file using Aspose.Cells | How much slower is saving an encrypted .xlsx with Aspose.Cells compared to unencrypted | Timing Aspose.Cells workbook.Save for large worksheets with encryption | Performance test for Aspose.Cells Xlsx save with password protection in .NET | Measure encryption overhead when writing large Excel workbooks with Aspose.Cells
// Tags: Aspose.Cells workbook save performance | password protection encryption Aspose.Cells .NET | benchmark large Excel workbook save time | measure Xlsx encryption overhead Aspose | performance testing Aspose.Cells save operation

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills a worksheet with 10,000 rows and 50 columns, records the time to save it as an unencrypted .xlsx, then applies a password for encryption and records the save time again, outputting both durations to demonstrate the performance impact of password‑protecting large Excel files with Aspose.Cells.
class Program
{
    static void Main()
    {
        // Initialize a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "LargeData";

        // Populate the worksheet with a large amount of data (e.g., 10000 rows x 50 columns)
        const int totalRows = 10000;
        const int totalCols = 50;
        for (int row = 0; row < totalRows; row++)
        {
            for (int col = 0; col < totalCols; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Measure time to save without encryption
        Stopwatch sw = new Stopwatch();
        string unencryptedPath = "UnencryptedWorkbook.xlsx";

        sw.Start();
        workbook.Save(unencryptedPath, SaveFormat.Xlsx);
        sw.Stop();
        Console.WriteLine($"Unencrypted save time: {sw.ElapsedMilliseconds} ms");

        // Apply password protection (encryption)
        workbook.Settings.Password = "StrongPassword123";

        // Measure time to save with encryption
        string encryptedPath = "EncryptedWorkbook.xlsx";

        sw.Restart();
        workbook.Save(encryptedPath, SaveFormat.Xlsx);
        sw.Stop();
        Console.WriteLine($"Encrypted save time: {sw.ElapsedMilliseconds} ms");
    }
}
