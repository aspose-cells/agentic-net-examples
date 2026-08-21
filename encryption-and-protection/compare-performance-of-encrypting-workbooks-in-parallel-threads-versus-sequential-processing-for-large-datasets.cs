// Title: Aspose.Cells .NET – Benchmark Sequential vs Parallel Workbook Encryption Performance
// Description: C# sample that creates 20 workbooks (5,000 × 10 cells each), applies a strong password with the StrongCryptographicProvider (128‑bit) and saves them first sequentially and then in parallel using Task.Run. The program measures and prints the elapsed time for both approaches, helping developers decide the fastest threading strategy for encrypting large Excel files with Aspose.Cells.
// Keywords: Aspose.Cells | C# | workbook encryption | benchmark | parallel processing | Task.Run | Parallel.ForEach | performance testing | large Excel files | encryption speed | multi‑threading | CPU utilization | memory usage
// Common Searches: Aspose.Cells encrypt multiple workbooks in parallel C# | benchmark workbook encryption speed Aspose.Cells .NET | sequential vs parallel Excel file encryption performance | how to speed up Aspose.Cells encryption with multithreading | measure encryption time for large Excel reports Aspose
// Developer Intent: Find out whether encrypting many large workbooks concurrently yields a measurable speed‑up compared with processing them one after another using Aspose.Cells for .NET.
// Use Cases: Batch‑generate and protect a set of high‑volume Excel reports before distribution. | Build a high‑throughput service that encrypts incoming workbooks concurrently to meet SLA requirements. | Profile encryption latency to choose the optimal threading model for a document‑processing pipeline.
// AI Prompts: Analyze the timing results from the sequential and parallel sections of the provided Aspose.Cells code and explain why a performance gap may exist. | Rewrite the parallel block using Parallel.ForEach with robust exception handling while keeping the same encryption options. | Add CPU and memory monitoring to the benchmark so that both execution time and resource consumption are captured for each workbook. | Suggest how to tune the degree of parallelism or switch encryption providers to further improve throughput on a multi‑core server.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace WorkbookEncryptionPerformance
{
    // C# sample that creates 20 workbooks (5,000 × 10 cells each), applies a strong password with the StrongCryptographicProvider (128‑bit) and saves them first sequentially and then in parallel using Task.Run. The program measures and prints the elapsed time for both approaches, helping developers decide the fastest threading strategy for encrypting large Excel files with Aspose.Cells.
    class Program
    {
        // Number of workbooks to process
        const int WorkbookCount = 20;
        // Size of each workbook (rows x columns)
        const int RowCount = 5000;
        const int ColumnCount = 10;
        // Encryption settings
        const string Password = "StrongPassword123";
        const EncryptionType Encryption = EncryptionType.StrongCryptographicProvider;
        const int KeyLength = 128;

        static void Main()
        {
            // Prepare output directory
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "EncryptedWorkbooks");
            Directory.CreateDirectory(outputDir);

            // ---------- Sequential Processing ----------
            Stopwatch swSequential = Stopwatch.StartNew();
            for (int i = 0; i < WorkbookCount; i++)
            {
                // Create a large workbook
                Workbook wb = CreateLargeWorkbook();

                // Apply encryption settings
                wb.Settings.Password = Password;
                wb.SetEncryptionOptions(Encryption, KeyLength);

                // Save encrypted workbook
                string filePath = Path.Combine(outputDir, $"Seq_{i + 1}.xlsx");
                wb.Save(filePath);
                wb.Dispose();
            }
            swSequential.Stop();
            Console.WriteLine($"Sequential processing time: {swSequential.ElapsedMilliseconds} ms");

            // ---------- Parallel Processing ----------
            Stopwatch swParallel = Stopwatch.StartNew();
            List<Task> tasks = new List<Task>();
            for (int i = 0; i < WorkbookCount; i++)
            {
                int index = i; // capture loop variable
                tasks.Add(Task.Run(() =>
                {
                    // Create a large workbook
                    Workbook wb = CreateLargeWorkbook();

                    // Apply encryption settings
                    wb.Settings.Password = Password;
                    wb.SetEncryptionOptions(Encryption, KeyLength);

                    // Save encrypted workbook
                    string filePath = Path.Combine(outputDir, $"Par_{index + 1}.xlsx");
                    wb.Save(filePath);
                    wb.Dispose();
                }));
            }
            Task.WaitAll(tasks.ToArray());
            swParallel.Stop();
            Console.WriteLine($"Parallel processing time: {swParallel.ElapsedMilliseconds} ms");
        }

        // Creates a workbook filled with dummy data (RowCount x ColumnCount)
        static Workbook CreateLargeWorkbook()
        {
            Workbook workbook = new Workbook(); // using the constructor rule
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            for (int row = 0; row < RowCount; row++)
            {
                for (int col = 0; col < ColumnCount; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }
            return workbook;
        }
    }
}
