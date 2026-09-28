// Title: Measure the time to save a password‑protected 100k‑row Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that creates a worksheet with 100,000 rows and 50 columns, applies a password via workbook.Settings.Password, saves the file, and prints the elapsed time for the Save call. | Extend the example to write the save duration (in seconds) to a log file and add an optional loop that saves the workbook both with and without a password to compare performance.
// Common Searches: asp.net benchmark save time of encrypted Excel file using Aspose.Cells | performance test for password protected workbook save in Aspose.Cells .NET | measure encryption overhead when saving large Excel workbook with Aspose.Cells | time taken to save 100000 rows Excel file with Aspose.Cells and password | compare Aspose.Cells save speed with and without workbook password
// Tags: Aspose.Cells save performance encryption .NET | password protected workbook save latency | large Excel workbook timing Aspose.Cells | benchmark workbook.Settings.Password impact | measure encryption overhead Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// // Creates a 100,000‑row by 50‑column worksheet, sets workbook.Settings.Password, saves as EncryptedLargeWorkbook.xlsx, and prints the elapsed milliseconds for the Save operation.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with a large amount of data (e.g., 100,000 rows × 50 columns)
        int totalRows = 100_000;
        int totalCols = 50;
        for (int row = 0; row < totalRows; row++)
        {
            for (int col = 0; col < totalCols; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Apply encryption by setting a password
        workbook.Settings.Password = "StrongPassword123";

        // Measure the time taken to save the encrypted workbook
        Stopwatch timer = Stopwatch.StartNew();
        workbook.Save("EncryptedLargeWorkbook.xlsx");
        timer.Stop();

        Console.WriteLine($"Save operation completed in {timer.ElapsedMilliseconds} ms.");
    }
}
