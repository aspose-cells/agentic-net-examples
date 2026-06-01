using System;
using System.Text;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsThreadSafetyDemo
{
    public class MultiThreadReadDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and obtain its cells collection
                Workbook wb = new Workbook();
                Cells cells = wb.Worksheets[0].Cells;

                // Enable multi‑thread reading support
                cells.MultiThreadReading = true;

                // Populate the first column with sample data (dates)
                int rowCount = 1000;
                for (int i = 0; i < rowCount; i++)
                {
                    cells[i, 0].PutValue(DateTime.Now.AddDays(i).ToOADate());
                }

                // Set up thread coordination variables
                int threadCount = 5;
                int cellsPerThread = rowCount / threadCount;
                int finished = 0;
                StringBuilder errors = new StringBuilder();

                // Launch multiple threads that read the same column concurrently
                for (int t = 0; t < threadCount; t++)
                {
                    int start = t * cellsPerThread;
                    int end = (t == threadCount - 1) ? rowCount : start + cellsPerThread;

                    Thread thread = new Thread(() =>
                    {
                        try
                        {
                            for (int r = start; r < end; r++)
                            {
                                // Read the cell value (no modification)
                                object value = cells[r, 0].Value;
                                Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{r},0] = {value}");
                            }
                            Interlocked.Increment(ref finished);
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

                // Wait for all threads to finish
                while (Volatile.Read(ref finished) < threadCount)
                {
                    Thread.Sleep(100);
                }

                // Report any errors encountered during reading
                if (errors.Length > 0)
                {
                    Console.WriteLine("Errors occurred during multi‑thread reading:");
                    Console.WriteLine(errors.ToString());
                }
                else
                {
                    Console.WriteLine("All threads completed successfully.");
                }

                // Save the workbook (demonstrates proper lifecycle usage)
                wb.Save("MultiThreadReadDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    // Entry point required for console application
    public class Program
    {
        public static void Main(string[] args)
        {
            MultiThreadReadDemo.Run();
        }
    }
}