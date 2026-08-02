// Title: Aspose.Cells C# Benchmark: Sequential vs Parallel Workbook Encryption for Large Datasets
// Description: A C# console app that generates five Excel workbooks with 200,000 rows each, applies password protection and encryption, then measures and compares the elapsed time when saving the files sequentially versus concurrently using the Task Parallel Library.
// Keywords: Aspose.Cells encryption benchmark | C# parallel workbook encryption | sequential vs parallel performance Aspose | large Excel file encryption .NET | measure encryption time Aspose.Cells
// Common Searches: Aspose.Cells encrypt multiple workbooks in parallel | benchmark Excel encryption performance C# | parallel processing for large Excel files Aspose | sequential encryption time Aspose.Cells | Task Parallel Library workbook encryption example
// Developer Intent: Find out whether encrypting several large Excel workbooks concurrently is faster than processing them one after another with Aspose.Cells.
// Use Cases: Batch creation of protected financial reports and selection of the optimal processing strategy. | Nightly generation of encrypted data exports where throughput matters. | Performance profiling of Aspose.Cells encryption in high‑volume data pipelines.
// AI Prompts: Generate a C# sample that encrypts 10 workbooks with 500,000 rows each, measuring both sequential and parallel execution times and summarizing the results. | Suggest ways to improve parallel encryption speed with Aspose.Cells, covering thread‑pool configuration and memory usage. | Show how to record per‑workbook encryption duration and aggregate statistics when using Task.WhenAll with Aspose.Cells.

using System;
using System.Diagnostics;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsEncryptionPerformance
{
    // A C# console app that generates five Excel workbooks with 200,000 rows each, applies password protection and encryption, then measures and compares the elapsed time when saving the files sequentially versus concurrently using the Task Parallel Library.
    class Program
    {
        // Number of workbooks to process
        private const int WorkbookCount = 5;

        // Number of rows per workbook (large dataset)
        private const int RowsPerWorkbook = 200_000;

        // Password to protect the workbook
        private const string Password = "StrongPassword123";

        // Encryption settings (ignored for 2007/2010 but required by API)
        private const EncryptionType Encryption = EncryptionType.StrongCryptographicProvider;
        private const int KeyLength = 128;

        // Creates a workbook, fills it with data, applies encryption, and saves to the specified file.
        private static void CreateAndEncryptWorkbook(string filePath)
        {
            // Create a new workbook (default Xlsx format)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate the worksheet with a large amount of data
            for (int row = 0; row < RowsPerWorkbook; row++)
            {
                // Example data: row number and a timestamp
                cells[row, 0].PutValue(row + 1);
                cells[row, 1].PutValue(DateTime.Now);
            }

            // Apply password protection
            workbook.Settings.Password = Password;

            // Apply encryption options (required by API, ignored for modern formats)
            workbook.SetEncryptionOptions(Encryption, KeyLength);

            // Save the encrypted workbook
            workbook.Save(filePath);
        }

        static void Main(string[] args)
        {
            // Prepare file names for sequential processing
            string[] sequentialFiles = new string[WorkbookCount];
            for (int i = 0; i < WorkbookCount; i++)
                sequentialFiles[i] = $"Sequential_{i + 1}.xlsx";

            // Measure sequential encryption time
            Stopwatch swSequential = Stopwatch.StartNew();
            for (int i = 0; i < WorkbookCount; i++)
            {
                CreateAndEncryptWorkbook(sequentialFiles[i]);
            }
            swSequential.Stop();
            Console.WriteLine($"Sequential processing time: {swSequential.ElapsedMilliseconds} ms");

            // Prepare file names for parallel processing
            string[] parallelFiles = new string[WorkbookCount];
            for (int i = 0; i < WorkbookCount; i++)
                parallelFiles[i] = $"Parallel_{i + 1}.xlsx";

            // Measure parallel encryption time using Task Parallel Library
            Stopwatch swParallel = Stopwatch.StartNew();
            Task[] tasks = new Task[WorkbookCount];
            for (int i = 0; i < WorkbookCount; i++)
            {
                int index = i; // Capture loop variable
                tasks[i] = Task.Run(() => CreateAndEncryptWorkbook(parallelFiles[index]));
            }
            Task.WaitAll(tasks);
            swParallel.Stop();
            Console.WriteLine($"Parallel processing time: {swParallel.ElapsedMilliseconds} ms");

            // Simple comparison output
            if (swParallel.ElapsedMilliseconds < swSequential.ElapsedMilliseconds)
                Console.WriteLine("Parallel encryption was faster.");
            else
                Console.WriteLine("Sequential encryption was faster or equal.");
        }
    }
}
