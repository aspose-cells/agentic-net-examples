using System;
using System.Text;
using System.Threading;
using Aspose.Cells;

class MultiThreadReadingDemo
{
    static void Main()
    {
        // Create a new workbook and obtain the Cells collection of the first worksheet
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Enable multi‑thread reading on the Cells collection
        cells.MultiThreadReading = true;

        // Populate the worksheet with sample data (e.g., 1000 rows × 10 columns)
        int totalRows = 1000;
        int totalCols = 10;
        for (int r = 0; r < totalRows; r++)
        {
            for (int c = 0; c < totalCols; c++)
            {
                cells[r, c].PutValue($"R{r}C{c}");
            }
        }

        // Define how many threads will read cells concurrently
        int threadCount = 5;
        // Define how many random reads each thread will perform
        int readsPerThread = 200;

        // Variables to track completion and collect any errors
        int finishedThreads = 0;
        StringBuilder errorLog = new StringBuilder();
        object errorLock = new object();

        // Launch the threads
        for (int i = 0; i < threadCount; i++)
        {
            Thread t = new Thread(() =>
            {
                try
                {
                    // Each thread uses its own Random instance to avoid contention
                    Random rnd = new Random(Guid.NewGuid().GetHashCode());

                    for (int j = 0; j < readsPerThread; j++)
                    {
                        int row = rnd.Next(totalRows);
                        int col = rnd.Next(totalCols);
                        object value = cells[row, col].Value; // Read cell value
                        Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{row},{col}] = {value}");
                    }

                    // Signal successful completion
                    Interlocked.Increment(ref finishedThreads);
                }
                catch (Exception ex)
                {
                    // Record any exception that occurs inside the thread
                    lock (errorLock)
                    {
                        errorLog.AppendLine($"Thread {Thread.CurrentThread.ManagedThreadId} error: {ex.Message}");
                    }
                }
            });

            t.IsBackground = true;
            t.Start();
        }

        // Wait until all threads have finished
        while (Volatile.Read(ref finishedThreads) < threadCount)
        {
            Thread.Sleep(100);
        }

        // Report results
        if (errorLog.Length > 0)
        {
            Console.WriteLine("Errors occurred during execution:");
            Console.WriteLine(errorLog.ToString());
        }
        else
        {
            Console.WriteLine("All threads completed successfully.");
        }

        // Save the workbook (demonstrates the required save operation)
        workbook.Save("MultiThreadReadingDemo.xlsx");
    }
}