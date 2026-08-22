// Title: Show data inconsistencies when reading Excel cells concurrently without enabling MultiThreadReading in Aspose.Cells for .NET
// AI Prompts: Write a C# console program that creates a workbook, fills column A with OADate dates, then launches several threads that each read the cell.Value of a range of rows without setting cells.MultiThreadReading, logging the thread ID and the retrieved value. | Enhance the multithreaded reading loop with try‑catch blocks, collect any exceptions in a thread‑safe StringBuilder, and display a summary of errors after all threads have completed.
// Common Searches: Aspose.Cells read cells from multiple threads without MultiThreadReading example | C# Aspose.Cells concurrent reading causing inconsistent values | how to reproduce thread safety issue in Aspose.Cells workbook reading | disable MultiThreadReading Aspose.Cells and observe data anomalies | multithreaded Excel cell access Aspose.Cells .NET tutorial
// Tags: concurrent cell reading Aspose.Cells | disable MultiThreadReading Aspose.Cells | thread safety Aspose.Cells workbook | raw cell value access parallel | C# multithreaded Excel read example

using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsMultiThreadDemo
{
    // The sample creates a new workbook, populates column A with OADate dates, and starts multiple threads that read the raw cell values concurrently while MultiThreadReading remains disabled. Each thread logs its ID and the value read, and any exceptions are captured and reported after all threads finish, illustrating potential data inconsistencies.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate the first column with sample data (dates)
            int totalRows = 1000;
            for (int i = 0; i < totalRows; i++)
            {
                // Store dates as OADate numbers to avoid formatting issues
                cells[i, 0].PutValue(DateTime.Now.AddDays(i).ToOADate());
            }

            // NOTE: We deliberately do NOT enable MultiThreadReading.
            // cells.MultiThreadReading = true; // <-- omitted on purpose

            // Prepare multithreading variables
            int threadCount = 5;
            int rowsPerThread = totalRows / threadCount;
            int[] finished = new int[1]; // shared counter
            StringBuilder errors = new StringBuilder();

            // Launch threads that read cell values concurrently
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
                            // Read the raw value (object) from the cell
                            object value = cells[row, 0].Value;

                            // Output the thread ID and the read value
                            Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{row},0] = {value}");
                        }

                        // Signal successful completion
                        Interlocked.Increment(ref finished[0]);
                    }
                    catch (Exception ex)
                    {
                        // Capture any exception that occurs during reading
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

            // Report any errors that were captured
            if (errors.Length > 0)
            {
                Console.WriteLine("Errors occurred during multithreaded reading:");
                Console.WriteLine(errors.ToString());
            }
            else
            {
                Console.WriteLine("All threads completed without throwing exceptions.");
                Console.WriteLine("Note: Even though no exception was thrown, reading without MultiThreadReading may produce inconsistent or unexpected values in real scenarios.");
            }
        }
    }
}
