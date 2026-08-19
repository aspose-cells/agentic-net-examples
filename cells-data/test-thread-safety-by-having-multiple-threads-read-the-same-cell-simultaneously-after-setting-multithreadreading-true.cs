// Title: Aspose.Cells .NET – Verify MultiThreadReading Thread‑Safety with Concurrent Cell Reads
// Description: Creates a workbook, turns on cells.MultiThreadReading, fills column A with 1,000 OADate values, then launches five threads that each read a separate row range from the same column. Completion is tracked with Interlocked and errors are collected in a StringBuilder, demonstrating that parallel reads succeed without exceptions.
// Keywords: Aspose.Cells | MultiThreadReading | thread safety | concurrent cell reading | C# | .NET | cells collection | Interlocked synchronization | multi‑threaded workbook access | performance testing
// Common Searches: enable MultiThreadReading in Aspose.Cells .NET | is Aspose.Cells thread‑safe for reading cells | sample code for concurrent cell reads with Aspose.Cells | how to use Interlocked with Aspose.Cells multi‑threading
// Developer Intent: Confirm that setting cells.MultiThreadReading = true allows multiple threads to read the same worksheet cells without raising errors.
// Use Cases: Validate thread‑safety when extracting data from large worksheets in a parallel processing pipeline. | Compare read throughput of a workbook with MultiThreadReading enabled versus a single‑threaded approach. | Implement error‑free concurrent data retrieval in a multi‑threaded API or background service.
// AI Prompts: Write a C# unit test that asserts no exception is thrown when five threads read the same column after enabling cells.MultiThreadReading. | Add timing logic using Stopwatch to compare the duration of concurrent reads against sequential reads in the sample. | Explain strategies for safely writing to cells while MultiThreadReading is active, including lock usage and write‑only sections.

using System;
using System.Text;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsThreadSafetyDemo
{
    // Creates a workbook, turns on cells.MultiThreadReading, fills column A with 1,000 OADate values, then launches five threads that each read a separate row range from the same column. Completion is tracked with Interlocked and errors are collected in a StringBuilder, demonstrating that parallel reads succeed without exceptions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Enable multi‑thread reading
            cells.MultiThreadReading = true;

            // Populate the first column with sample data (dates converted to OADate)
            int totalRows = 1000;
            for (int i = 0; i < totalRows; i++)
            {
                cells[i, 0].PutValue(DateTime.Now.AddDays(i).ToOADate());
            }

            // Prepare thread synchronization helpers
            int threadCount = 5;
            int rowsPerThread = totalRows / threadCount;
            int[] finished = new int[1];               // used with Interlocked
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
                            // Read the cell value (no formatting APIs are used)
                            object value = cells[row, 0].Value;
                            Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{row},0] = {value}");
                        }

                        // Signal successful completion
                        Interlocked.Increment(ref finished[0]);
                    }
                    catch (Exception ex)
                    {
                        // Capture any unexpected exception
                        lock (errors)
                        {
                            errors.AppendLine($"Thread {Thread.CurrentThread.ManagedThreadId} error: {ex.Message}");
                        }
                    }
                });

                thread.Start();
            }

            // Wait until all threads have reported completion
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
        }
    }
}
