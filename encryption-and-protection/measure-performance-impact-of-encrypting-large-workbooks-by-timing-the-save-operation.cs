// Title: Measure the time required to save a password‑protected large Excel workbook with Aspose.Cells for .NET
// Common Searches: how long does it take to save a password protected Excel file using Aspose.Cells in C# | benchmarking save performance of a large encrypted XLSX workbook with Aspose.Cells | measure encryption overhead when saving a 10,000 row workbook with Aspose.Cells | C# Aspose.Cells timing for saving an encrypted workbook to disk | performance impact of Workbook.Settings.Password on save speed in Aspose.Cells
// Tags: Aspose.Cells workbook save performance | password protected XLSX save latency | large workbook encryption overhead Aspose.Cells | C# Stopwatch timing Aspose.Cells save | encrypt Excel file with Workbook.Settings.Password

using System;
using System.Diagnostics;
using Aspose.Cells;
using System.IO;

// The example creates a 10,000‑row by 50‑column workbook, fills it with sample data, applies a password via Workbook.Settings.Password, uses a Stopwatch to time the Save call, writes the encrypted file as XLSX, and outputs the elapsed milliseconds.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define size of the large workbook (e.g., 10,000 rows x 50 columns)
        int totalRows = 10000;
        int totalCols = 50;

        // Fill the workbook with sample data
        for (int row = 0; row < totalRows; row++)
        {
            for (int col = 0; col < totalCols; col++)
            {
                cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Apply encryption by setting a password
        workbook.Settings.Password = "MySecretPassword";

        // Prepare timing
        Stopwatch sw = new Stopwatch();
        sw.Start();

        // Save the encrypted workbook to disk
        string outputPath = "EncryptedLargeWorkbook.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        // Stop timing
        sw.Stop();

        // Output the elapsed time
        Console.WriteLine($"Saving encrypted workbook took {sw.ElapsedMilliseconds} ms");
    }
}
