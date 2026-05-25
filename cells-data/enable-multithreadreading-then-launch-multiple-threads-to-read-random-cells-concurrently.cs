using System;
using System.Text;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsMultiThreadReadingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate the worksheet with sample data (1000 rows, 1 column)
            int totalRows = 1000;
            for (int i = 0; i < totalRows; i++)
            {
                // Store a simple integer value in each cell
                cells[i, 0].PutValue(i);
            }

            // Enable multi‑thread reading on the cells collection
            cells.MultiThreadReading = true;

            // Number of threads that will read cells concurrently
            int threadCount = 5;
            // Each thread will perform this many read operations
            int readsPerThread = 200;

            // Variables to track completion and errors
            int completedThreads = 0;
            StringBuilder errorLog = new StringBuilder();

            // Launch the threads
            for (int t = 0; t < threadCount; t++)
            {
                Thread thread = new Thread(() =>
                {
                    try
                    {
                        // Each thread uses its own Random instance to avoid contention
                        Random rnd = new Random(Thread.CurrentThread.ManagedThreadId);
                        for (int r = 0; r < readsPerThread; r++)
                        {
                            // Choose a random row index within the populated range
                            int rowIndex = rnd.Next(0, totalRows);
                            // Read the cell value (thread‑safe because MultiThreadReading is true)
                            object value = cells[rowIndex, 0].Value;
                            Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{rowIndex},0] = {value}");
                        }

                        // Signal successful completion of this thread
                        Interlocked.Increment(ref completedThreads);
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
            while (Volatile.Read(ref completedThreads) < threadCount)
            {
                Thread.Sleep(100);
            }

            // Report any errors that were captured
            if (errorLog.Length > 0)
            {
                Console.WriteLine("Errors occurred during multi‑thread reading:");
                Console.WriteLine(errorLog.ToString());
            }
            else
            {
                Console.WriteLine("All threads completed successfully without errors.");
            }

            // (Optional) Save the workbook to verify that data remains intact
            workbook.Save("MultiThreadReadingResult.xlsx");
        }
    }
}