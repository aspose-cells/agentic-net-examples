using System;
using System.Text;
using System.Threading;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: use constructor)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Enable multi‑thread reading (required for safe concurrent access)
        cells.MultiThreadReading = true;

        // Populate the first column with sample data
        int rowCount = 1000;
        for (int i = 0; i < rowCount; i++)
        {
            // Store dates as OLE Automation numbers for simplicity
            cells[i, 0].PutValue(DateTime.Now.AddDays(i).ToOADate());
        }

        // Set up multi‑threaded reading
        int threadCount = 5;
        int cellsPerThread = rowCount / threadCount;
        int[] finished = new int[1];               // shared counter for completed threads
        StringBuilder errors = new StringBuilder(); // collect any thread errors

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
                        // Read cell value (thread‑safe because MultiThreadReading is true)
                        object value = cells[r, 0].Value;
                        Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Row {r} = {value}");
                    }
                    Interlocked.Increment(ref finished[0]); // signal completion
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

        // Wait until all threads have finished
        while (finished[0] < threadCount)
        {
            Thread.Sleep(200);
        }

        // Report any errors
        if (errors.Length > 0)
        {
            Console.WriteLine("Errors occurred during reading:");
            Console.WriteLine(errors.ToString());
        }
        else
        {
            Console.WriteLine("All threads completed successfully.");
        }

        // Save the workbook (lifecycle rule: use Save method)
        workbook.Save("ConcurrentReadDemo.xlsx");
    }
}