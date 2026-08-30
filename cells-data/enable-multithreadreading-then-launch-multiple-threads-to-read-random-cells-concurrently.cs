// Title: Enable MultiThreadReading and perform concurrent random cell reads with multiple threads in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that sets Cells.MultiThreadReading = true, populates a worksheet with sample data, and launches several threads that each read random cell values safely. | Generate a multi‑threaded Aspose.Cells example that uses ThreadLocal<Random> and Interlocked to read 200 random cells per thread without race conditions.
// Common Searches: Aspose.Cells C# enable MultiThreadReading for thread‑safe cell access | example of reading random Excel cells concurrently using multiple threads in .NET | how to use ThreadLocal Random with Aspose.Cells multi‑threaded reading | C# multi‑threaded worksheet reading Aspose.Cells sample code | prevent race conditions when reading cells with Aspose.Cells MultiThreadReading
// Tags: Aspose.Cells MultiThreadReading activation | C# concurrent cell reading Aspose.Cells | Excel worksheet random cell access multi-threaded | multi-threaded cell read synchronization Aspose.Cells | launch multiple threads for worksheet reading Aspose.Cells

using System;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsMultiThreadReadingDemo
{
    // The program creates a new workbook, enables MultiThreadReading on its Cells collection, fills 1,000 rows with integer values, and starts five threads. Each thread reads 200 random cells from column A, prints the values, and signals completion using Interlocked. Errors are captured in a thread‑safe StringBuilder, and a final success or error report is displayed after all threads finish.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Enable multi‑thread reading on the cells collection
            cells.MultiThreadReading = true;

            // Populate the worksheet with sample data (1000 rows, 1 column)
            int totalRows = 1000;
            for (int i = 0; i < totalRows; i++)
            {
                // Store a simple integer value in each cell
                cells[i, 0].PutValue(i);
            }

            // Number of concurrent threads that will read cells
            int threadCount = 5;
            // Counter to track completed threads
            int finishedThreads = 0;
            // StringBuilder to collect any errors from threads
            System.Text.StringBuilder errorLog = new System.Text.StringBuilder();

            // Random number generator (thread‑safe via ThreadLocal)
            ThreadLocal<Random> threadRandom = new ThreadLocal<Random>(() => new Random());

            // Launch the threads
            for (int t = 0; t < threadCount; t++)
            {
                Thread thread = new Thread(() =>
                {
                    try
                    {
                        // Each thread reads 200 random cells
                        for (int i = 0; i < 200; i++)
                        {
                            int rowIndex = threadRandom.Value.Next(0, totalRows);
                            object value = cells[rowIndex, 0].Value; // Safe read because MultiThreadReading = true
                            Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{rowIndex},0] = {value}");
                        }

                        // Signal successful completion
                        Interlocked.Increment(ref finishedThreads);
                    }
                    catch (Exception ex)
                    {
                        // Record any exception that occurs during reading
                        lock (errorLog)
                        {
                            errorLog.AppendLine($"Thread {Thread.CurrentThread.ManagedThreadId} error: {ex.Message}");
                        }
                    }
                });

                thread.Start();
            }

            // Wait until all threads have finished
            while (Volatile.Read(ref finishedThreads) < threadCount)
            {
                Thread.Sleep(100);
            }

            // Report any errors or success
            if (errorLog.Length > 0)
            {
                Console.WriteLine("Errors occurred during multi‑thread reading:");
                Console.WriteLine(errorLog.ToString());
            }
            else
            {
                Console.WriteLine("All threads completed successfully.");
            }
        }
    }
}
