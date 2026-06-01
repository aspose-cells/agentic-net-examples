using System;
using System.Collections.Generic;
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

            // Populate the first column with sample data (dates as OADate numbers)
            int totalRows = 1000;
            for (int i = 0; i < totalRows; i++)
            {
                cells[i, 0].PutValue(DateTime.Now.AddDays(i).ToOADate());
            }

            // NOTE: We deliberately DO NOT enable MultiThreadReading.
            // cells.MultiThreadReading = true; // <-- not set

            // Prepare multithreaded reading
            int threadCount = 5;
            int rowsPerThread = totalRows / threadCount;
            int[] finished = new int[1]; // shared counter
            StringBuilder errorLog = new StringBuilder();

            // Start multiple threads that read cell values concurrently
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
                            // Read the cell value without thread‑safety guarantees
                            object value = cells[row, 0].Value;
                            Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{row},0] = {value}");
                        }

                        // Signal successful completion
                        Interlocked.Increment(ref finished[0]);
                    }
                    catch (Exception ex)
                    {
                        // Capture any unexpected exception (e.g., data inconsistency)
                        lock (errorLog)
                        {
                            errorLog.AppendLine($"Thread {Thread.CurrentThread.ManagedThreadId} error: {ex.Message}");
                        }
                    }
                });

                thread.Start();
            }

            // Wait for all threads to finish
            while (finished[0] < threadCount)
            {
                Thread.Sleep(200);
            }

            // Report any errors that occurred during reading
            if (errorLog.Length > 0)
            {
                Console.WriteLine("Errors detected during multithreaded reading (MultiThreadReading disabled):");
                Console.WriteLine(errorLog.ToString());
            }
            else
            {
                Console.WriteLine("All threads completed without throwing exceptions, but data may still be inconsistent.");
            }

            // Save the workbook (optional, just to demonstrate normal save flow)
            workbook.Save("MultiThreadReadingDemo.xlsx");
        }
    }
}