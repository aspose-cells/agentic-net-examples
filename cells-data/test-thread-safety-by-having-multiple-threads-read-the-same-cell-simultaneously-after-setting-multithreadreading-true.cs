// Title: Testing Aspose.Cells thread safety by concurrently reading the same column with MultiThreadReading enabled in C#
// AI Prompts: Enable MultiThreadReading on a Cells collection, populate a column with OADate values, and launch several threads that each read a distinct range of cells to verify that no race conditions occur. | Apply Interlocked to count finished threads and use a StringBuilder to collect any exceptions while performing concurrent cell reads in Aspose.Cells. | Extend the sample to read multiple columns in parallel, logging each thread's ID and the retrieved values for deeper thread‑safety validation.
// Common Searches: Aspose.Cells C# example for multi‑threaded cell reading with MultiThreadReading property | how to verify thread safety of Cells collection in Aspose.Cells | concurrent read of the same column using Aspose.Cells and C# threads | C# Aspose.Cells multi‑thread reading without exceptions | test Aspose.Cells MultiThreadReading performance with multiple threads
// Tags: Aspose.Cells MultiThreadReading concurrency test | C# concurrent cell value retrieval | thread‑safe cell reading Aspose.Cells | Interlocked thread completion tracking | StringBuilder exception aggregation in multi‑threaded Aspose.Cells

using System;
using System.Text;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsThreadSafetyDemo
{
    // The program creates a new Workbook, turns on MultiThreadReading for its Cells collection, fills the first column with 1,000 OADate values, and starts five threads that each read a separate segment of that column simultaneously. It uses Interlocked to count completed threads and a StringBuilder to capture any exceptions, finally reporting whether all threads finished without errors, thereby demonstrating thread‑safe concurrent reads.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Enable multi‑thread reading on the cells collection
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
                            // Read the cell value; this is the operation we are testing for thread safety
                            object value = cells[row, 0].Value;
                            Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{row},0] = {value}");
                        }

                        // Signal successful completion of this thread
                        Interlocked.Increment(ref finished[0]);
                    }
                    catch (Exception ex)
                    {
                        // Capture any exception details
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
