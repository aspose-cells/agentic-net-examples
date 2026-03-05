using System;
using System.Text;
using System.Threading;
using Aspose.Cells;
using Range = Aspose.Cells.Range;

namespace AsposeCellsMultiThreadReadDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Enable multi‑thread reading on the cells collection
            cells.MultiThreadReading = true;

            // Populate the first column with sample date values
            int rowCount = 1000;
            for (int i = 0; i < rowCount; i++)
            {
                // Store dates as OADate numbers (Excel internal date format)
                cells[i, 0].PutValue(DateTime.Now.AddDays(i).ToOADate());
            }

            // Define threading parameters
            int threadCount = 5;
            int cellsPerThread = rowCount / threadCount;
            int[] finishedCount = new int[1]; // shared counter for completed threads
            StringBuilder errors = new StringBuilder();

            // Launch multiple threads to read cell values concurrently
            for (int i = 0; i < threadCount; i++)
            {
                int start = i * cellsPerThread;
                int end = (i == threadCount - 1) ? rowCount : start + cellsPerThread;

                Thread t = new Thread(() =>
                {
                    try
                    {
                        for (int row = start; row < end; row++)
                        {
                            // Read the raw value (object) from the cell
                            object value = cells[row, 0].Value;
                            Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{row},0] = {value}");
                        }

                        // Signal successful completion of this thread
                        Interlocked.Increment(ref finishedCount[0]);
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

                t.Start();
            }

            // Wait until all threads have finished
            while (finishedCount[0] < threadCount)
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
                Console.WriteLine("All threads completed successfully.");
            }

            // Optionally save the workbook (demonstrates normal save lifecycle)
            workbook.Save("MultiThreadReadResult.xlsx");
        }
    }
}