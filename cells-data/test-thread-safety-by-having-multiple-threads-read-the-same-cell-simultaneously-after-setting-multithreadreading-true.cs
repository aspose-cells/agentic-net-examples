// Title: Aspose.Cells C# MultiThreadReading Demo – Safe Concurrent Cell Reads
// Description: A complete C# example that creates a workbook, activates Cells.MultiThreadReading, fills the first column with OADate values, and launches several threads that read distinct row ranges simultaneously. The code uses Interlocked for completion tracking, captures any exceptions, reports the outcome, and saves the workbook, illustrating thread‑safe read‑only access in Aspose.Cells.
// Keywords: Aspose.Cells | MultiThreadReading | thread safety | C# concurrent cell read | parallel Excel access | Aspose.Cells .NET example | multi‑thread reading workbook | Excel cell thread safety | Aspose.Cells performance | cell reading multithread
// Common Searches: enable MultiThreadReading Aspose.Cells C# | C# read Excel cells from multiple threads | Aspose.Cells thread safety example | test concurrent cell reads Aspose.Cells | MultiThreadReading performance Aspose.Cells
// Developer Intent: Confirm that setting Cells.MultiThreadReading = true allows multiple threads to read the same cells concurrently without throwing exceptions.
// Use Cases: Validate read‑only thread safety before processing large worksheets in parallel. | Benchmark multi‑threaded cell‑read performance versus single‑threaded execution. | Integrate safe concurrent Excel data extraction into high‑throughput services. | Create unit tests that ensure Aspose.Cells read operations are race‑condition free. | Demonstrate proper synchronization (Interlocked, StringBuilder) when verifying thread safety.
// AI Prompts: Write a C# unit test that sets Cells.MultiThreadReading = true, spawns N threads to read a shared range, and asserts no exceptions occur. | Explain how Aspose.Cells implements thread‑safe reading when MultiThreadReading is enabled and list any configuration limits. | Suggest modifications to measure per‑thread read latency and detect subtle race conditions during concurrent cell access.

using System;
using System.Text;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsThreadSafetyDemo
{
    // A complete C# example that creates a workbook, activates Cells.MultiThreadReading, fills the first column with OADate values, and launches several threads that read distinct row ranges simultaneously. The code uses Interlocked for completion tracking, captures any exceptions, reports the outcome, and saves the workbook, illustrating thread‑safe read‑only access in Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Enable multi‑thread reading on the cells collection
            cells.MultiThreadReading = true;

            // Populate the first column with sample data (dates stored as OADate)
            int totalRows = 1000;
            for (int i = 0; i < totalRows; i++)
            {
                cells[i, 0].PutValue(DateTime.Now.AddDays(i).ToOADate());
            }

            // Prepare thread‑synchronization helpers
            int threadCount = 5;
            int rowsPerThread = totalRows / threadCount;
            int[] finished = new int[1];                     // used with Interlocked
            StringBuilder errors = new StringBuilder();

            // Launch multiple threads that read the same column concurrently
            for (int t = 0; t < threadCount; t++)
            {
                int startRow = t * rowsPerThread;
                int endRow = (t == threadCount - 1) ? totalRows : startRow + rowsPerThread;

                Thread thread = new Thread(() =>
                {
                    try
                    {
                        for (int row = startRow; row < endRow; row++)
                        {
                            // Read the cell value; no modification is performed
                            object value = cells[row, 0].Value;
                            Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{row},0] = {value}");
                        }
                        Interlocked.Increment(ref finished[0]);
                    }
                    catch (Exception ex)
                    {
                        lock (errors)
                        {
                            errors.AppendLine($"Thread {Thread.CurrentThread.ManagedThreadId} error: {ex.Message}");
                        }
                    }
                });

                thread.Start();
            }

            // Wait until all threads have signaled completion
            while (finished[0] < threadCount)
            {
                Thread.Sleep(200);
            }

            // Report any errors that occurred during reading
            if (errors.Length > 0)
            {
                Console.WriteLine("Errors occurred during multi‑thread reading:");
                Console.WriteLine(errors.ToString());
            }
            else
            {
                Console.WriteLine("All threads completed successfully without errors.");
            }

            // Optionally save the workbook (demonstrates normal lifecycle usage)
            workbook.Save("ThreadSafetyDemo.xlsx");
        }
    }
}
