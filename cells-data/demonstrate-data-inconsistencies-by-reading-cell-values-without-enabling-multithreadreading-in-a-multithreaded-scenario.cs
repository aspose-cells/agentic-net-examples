using System;
using System.Text;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsMultiThreadDemo
{
    class Program
    {
        // Number of rows to populate and read
        private const int RowCount = 1000;
        // Number of concurrent threads
        private const int ThreadCount = 5;

        static void Main()
        {
            // -------------------------------------------------
            // Create a new workbook and get the first worksheet
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // Populate the first column with sequential dates
            // -------------------------------------------------
            for (int i = 0; i < RowCount; i++)
            {
                // Store as OADate to keep numeric representation
                cells[i, 0].PutValue(DateTime.Now.AddDays(i).ToOADate());
            }

            // -------------------------------------------------
            // Do NOT enable MultiThreadReading (default is false)
            // -------------------------------------------------
            // cells.MultiThreadReading = true; // <-- intentionally omitted

            // -------------------------------------------------
            // Prepare synchronization structures
            // -------------------------------------------------
            int[] finished = new int[1];               // will be updated atomically
            StringBuilder errors = new StringBuilder(); // collect any exceptions

            // -------------------------------------------------
            // Launch multiple threads that read the same range
            // -------------------------------------------------
            int rowsPerThread = RowCount / ThreadCount;
            for (int t = 0; t < ThreadCount; t++)
            {
                int start = t * rowsPerThread;
                int end = (t == ThreadCount - 1) ? RowCount : start + rowsPerThread;

                Thread thread = new Thread(() =>
                {
                    try
                    {
                        for (int r = start; r < end; r++)
                        {
                            // Read the cell value (may be inconsistent without MultiThreadReading)
                            object value = cells[r, 0].Value;

                            // Output to console for visual inspection
                            Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{r},0] = {value}");
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

            // -------------------------------------------------
            // Wait for all threads to finish
            // -------------------------------------------------
            while (finished[0] < ThreadCount)
            {
                Thread.Sleep(200);
            }

            // -------------------------------------------------
            // Report any errors captured during reading
            // -------------------------------------------------
            if (errors.Length > 0)
            {
                Console.WriteLine("Errors occurred during multi‑threaded reading:");
                Console.WriteLine(errors.ToString());
            }
            else
            {
                Console.WriteLine("All threads completed without throwing exceptions.");
                Console.WriteLine("Note: Because MultiThreadReading was not enabled, the read values may be inconsistent or stale.");
            }

            // -------------------------------------------------
            // Save the workbook (optional, demonstrates normal save flow)
            // -------------------------------------------------
            workbook.Save("MultiThreadInconsistentDemo.xlsx");
        }
    }
}