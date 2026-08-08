// Title: C# Example: Enable Aspose.Cells MultiThreadReading for Parallel Random Cell Access
// Description: Demonstrates how to turn on the Cells.MultiThreadReading flag, fill a worksheet with 1,000 rows, and launch five threads that each read 200 random cells safely. The code uses lock‑protected Random, Interlocked counters, and error aggregation, then reports the outcome and saves the workbook.
// Keywords: Aspose.Cells | MultiThreadReading | parallel cell reading | C# multithreading | thread‑safe random access | Aspose.Cells .NET example | concurrent workbook read | performance optimization | Excel data extraction | GitHub sample
// Common Searches: Aspose.Cells enable MultiThreadReading C# | read Excel cells concurrently with Aspose.Cells | thread‑safe random cell access Aspose.Cells | parallel Excel data read .NET | sample code for Aspose.Cells multithreading
// Developer Intent: Activate Aspose.Cells multi‑thread reading and retrieve random cell values concurrently from a workbook.
// Use Cases: Speed up large‑scale data extraction by reading many cells in parallel. | Validate the stability of Aspose.Cells under concurrent read workloads. | Integrate parallel cell reads into reporting or analytics pipelines.
// AI Prompts: Write C# code that uses Aspose.Cells MultiThreadReading to read cells in parallel and aggregate the results. | Explain best practices for synchronizing Random and handling exceptions when reading cells concurrently with Aspose.Cells. | Show how to extend the example to perform parallel writes while maintaining thread safety.

using System;
using System.Text;
using System.Threading;
using Aspose.Cells;

// Demonstrates how to turn on the Cells.MultiThreadReading flag, fill a worksheet with 1,000 rows, and launch five threads that each read 200 random cells safely. The code uses lock‑protected Random, Interlocked counters, and error aggregation, then reports the outcome and saves the workbook.
class MultiThreadReadDemo
{
    static void Main()
    {
        // Create a new workbook and get the cells collection
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Enable multi‑thread reading on the cells model
        cells.MultiThreadReading = true;

        // Populate the first column with sample data
        int totalRows = 1000;
        for (int i = 0; i < totalRows; i++)
        {
            cells[i, 0].PutValue($"Row {i}");
        }

        // Number of concurrent threads
        int threadCount = 5;
        // Counter to track completed threads
        int[] finished = new int[1];
        // Collect any errors that occur in threads
        StringBuilder errors = new StringBuilder();

        // Random number generator (shared, but thread‑safe usage via lock)
        Random rnd = new Random();

        for (int i = 0; i < threadCount; i++)
        {
            Thread t = new Thread(() =>
            {
                try
                {
                    // Each thread reads 200 random cells from the column
                    for (int j = 0; j < 200; j++)
                    {
                        int row;
                        // Ensure thread‑safe random number generation
                        lock (rnd)
                        {
                            row = rnd.Next(totalRows);
                        }

                        object value = cells[row, 0].Value;
                        Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{row},0] = {value}");
                    }

                    // Signal successful completion
                    Interlocked.Increment(ref finished[0]);
                }
                catch (Exception ex)
                {
                    // Record any exception details
                    lock (errors)
                    {
                        errors.AppendLine($"Thread {Thread.CurrentThread.ManagedThreadId} error: {ex.Message}");
                    }
                }
            });

            t.Start();
        }

        // Wait until all threads have finished
        while (finished[0] < threadCount)
        {
            Thread.Sleep(100);
        }

        // Report results
        if (errors.Length > 0)
        {
            Console.WriteLine("Errors occurred:");
            Console.WriteLine(errors.ToString());
        }
        else
        {
            Console.WriteLine("All threads completed successfully");
        }

        // Optional: save the workbook to verify data integrity
        workbook.Save("MultiThreadReadDemo.xlsx");
    }
}
