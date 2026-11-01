// Title: Benchmark sequential vs parallel encryption of large Excel workbooks using Aspose.Cells for .NET
// Common Searches: C# how to encrypt many large Excel files concurrently with Aspose.Cells | measure performance of Aspose.Cells password protection for big workbooks | sequential versus parallel encryption speed using Aspose.Cells .NET | benchmark Aspose.Cells workbook encryption on a 100k‑row dataset | parallel processing of Excel file encryption with Aspose.Cells example
// Tags: parallel workbook encryption Aspose.Cells | sequential Excel password protection Aspose.Cells | Aspose.Cells encryption performance benchmark | large dataset workbook generation C# | Parallel.ForEach Aspose.Cells encryption

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample creates ten Excel workbooks each with 100,000 rows and 10 columns, applies a password, and records the elapsed time for encrypting the files sequentially and then in parallel using Parallel.ForEach, demonstrating the performance impact of concurrent encryption with Aspose.Cells.
class Program
{
    // Number of workbooks to process
    const int WorkbookCount = 10;
    // Size of each workbook (rows x columns)
    const int RowCount = 100_000;
    const int ColumnCount = 10;
    // Encryption password
    const string Password = "Secret123";

    static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Path.GetTempPath(), "AsposeEncryptDemo");
        Directory.CreateDirectory(outputDir);

        // Generate list of file paths for the workbooks
        List<string> filePaths = new List<string>();
        for (int i = 0; i < WorkbookCount; i++)
        {
            string filePath = Path.Combine(outputDir, $"Workbook_{i}.xlsx");
            filePaths.Add(filePath);
        }

        // Sequential encryption
        Stopwatch swSequential = Stopwatch.StartNew();
        foreach (string path in filePaths)
        {
            // Create workbook with large dataset
            Workbook wb = CreateLargeWorkbook();

            // Apply password protection
            wb.Settings.Password = Password;

            // Save encrypted workbook
            wb.Save(path, SaveFormat.Xlsx);
        }
        swSequential.Stop();
        Console.WriteLine($"Sequential encryption time: {swSequential.ElapsedMilliseconds} ms");

        // Clean up files before parallel test
        foreach (string path in filePaths)
        {
            if (File.Exists(path))
                File.Delete(path);
        }

        // Parallel encryption
        Stopwatch swParallel = Stopwatch.StartNew();
        Parallel.ForEach(filePaths, path =>
        {
            // Create workbook with large dataset
            Workbook wb = CreateLargeWorkbook();

            // Apply password protection
            wb.Settings.Password = Password;

            // Save encrypted workbook
            wb.Save(path, SaveFormat.Xlsx);
        });
        swParallel.Stop();
        Console.WriteLine($"Parallel encryption time: {swParallel.ElapsedMilliseconds} ms");
    }

    // Creates a workbook filled with a large dataset (RowCount x ColumnCount)
    static Workbook CreateLargeWorkbook()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate cells with sample data
        for (int row = 0; row < RowCount; row++)
        {
            for (int col = 0; col < ColumnCount; col++)
            {
                cells[row, col].PutValue($"R{row}_C{col}");
            }
        }

        return workbook;
    }
}
