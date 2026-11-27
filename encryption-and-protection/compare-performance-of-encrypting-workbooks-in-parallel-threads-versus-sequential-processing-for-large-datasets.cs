// Title: Measuring sequential versus parallel workbook encryption performance with Aspose.Cells for .NET
// Common Searches: how to measure Aspose.Cells workbook encryption speed in C# | compare parallel and sequential encryption of large Excel files using Aspose.Cells | benchmark Aspose.Cells password protection for worksheets with 100k rows in .NET | performance impact of Parallel.For when encrypting multiple workbooks with Aspose.Cells | C# timing test for encrypting many Excel workbooks with Aspose.Cells
// Tags: Aspose.Cells workbook encryption performance | parallel workbook protection Aspose.Cells .NET | sequential Excel file encryption benchmark | large dataset Excel encryption C# | Parallel.For Aspose.Cells encryption timing | memory stream workbook save Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample creates 20 workbooks each containing 100,000 rows and 10 columns, applies a password, saves them to a memory stream, and records the elapsed time for sequential processing versus Parallel.For multithreaded processing, allowing developers to compare encryption throughput.
class Program
{
    // Number of workbooks to encrypt for the test
    const int WorkbookCount = 20;

    // Size of each workbook (rows x columns)
    const int RowCount = 100_000;
    const int ColumnCount = 10;

    // Password used for encryption
    const string Password = "Secret123";

    static void Main()
    {
        try
        {
            // Warm up Aspose.Cells (first call may include JIT overhead)
            CreateAndEncryptWorkbook();

            // Sequential processing
            var swSequential = Stopwatch.StartNew();
            for (int i = 0; i < WorkbookCount; i++)
            {
                CreateAndEncryptWorkbook();
            }
            swSequential.Stop();

            // Parallel processing
            var swParallel = Stopwatch.StartNew();
            Parallel.For(0, WorkbookCount, i =>
            {
                CreateAndEncryptWorkbook();
            });
            swParallel.Stop();

            Console.WriteLine($"Sequential encryption time: {swSequential.ElapsedMilliseconds} ms");
            Console.WriteLine($"Parallel encryption time:   {swParallel.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Creates a workbook with a large dataset, encrypts it, and saves to a memory stream.
    static void CreateAndEncryptWorkbook()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Set workbook password for encryption
            workbook.Settings.Password = Password;

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;

            // Populate the worksheet with dummy data
            for (int row = 0; row < RowCount; row++)
            {
                for (int col = 0; col < ColumnCount; col++)
                {
                    cells[row, col].PutValue($"R{row}_C{col}");
                }
            }

            // Save to a memory stream to avoid disk I/O impact on timing
            using (var ms = new MemoryStream())
            {
                // Use default OoxmlSaveOptions (no password needed here)
                var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
                workbook.Save(ms, saveOptions);
                // The encrypted workbook is now in 'ms'. No further action needed for the benchmark.
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during workbook creation/encryption: {ex.Message}");
        }
    }
}
