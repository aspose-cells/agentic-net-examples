// Title: Demonstrate cell value inconsistencies in multithreaded reads without MultiThreadReading using Aspose.Cells for .NET
// Common Searches: Aspose.Cells multithreaded read example without MultiThreadReading in C# | how to detect race conditions when reading Excel cells concurrently with Aspose.Cells | C# code to reproduce data inconsistencies using ThreadPool and Aspose.Cells | reading Excel workbook in parallel without enabling MultiThreadReading Aspose.Cells | sample program showing inconsistent cell values in concurrent reads Aspose.Cells .NET
// Tags: Aspose.Cells multithreaded read without MultiThreadReading | C# ThreadPool random Excel cell reads | parallel Excel Xlsx read race condition | LoadOptions read‑only workbook concurrent access | detect inconsistent cell values Aspose.Cells

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

// The example creates a workbook with deterministic values (0‑999) in column A, saves it to a memory stream, and reloads it with LoadOptions. Eight ThreadPool threads each perform 10,000 random reads of cells in the first worksheet without enabling MultiThreadReading. Each read validates that the cell value matches the row index, incrementing an inconsistency counter on mismatch or exception. After all threads finish, the total reads and any detected inconsistencies are printed, illustrating potential data races when reading cells concurrently without MultiThreadReading.
class Program
{
    static void Main()
    {
        try
        {
            // -------------------------------------------------
            // Create a workbook with deterministic data
            // -------------------------------------------------
            Workbook wb = new Workbook();                     // create workbook
            Worksheet ws = wb.Worksheets[0];
            for (int i = 0; i < 1000; i++)
            {
                ws.Cells[i, 0].PutValue(i);                  // cell A1..A1000 = 0..999
            }

            // -------------------------------------------------
            // Save to a memory stream (no physical file needed)
            // -------------------------------------------------
            using (MemoryStream ms = new MemoryStream())
            {
                wb.Save(ms, SaveFormat.Xlsx);                // save workbook
                ms.Position = 0;                             // reset stream position

                // -------------------------------------------------
                // Load workbook (thread‑safe reading is inherent for read‑only operations)
                // -------------------------------------------------
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                Workbook readWb = new Workbook(ms, loadOptions);
                Worksheet readWs = readWb.Worksheets[0];

                // -------------------------------------------------
                // Multithreaded read scenario
                // -------------------------------------------------
                int threadCount = 8;                         // number of concurrent threads
                int readsPerThread = 10000;                  // how many reads each thread performs
                int inconsistencies = 0;                     // counter for mismatched values
                ManualResetEventSlim done = new ManualResetEventSlim(false);
                int completedThreads = 0;

                for (int t = 0; t < threadCount; t++)
                {
                    ThreadPool.QueueUserWorkItem(_ =>
                    {
                        try
                        {
                            // Use a thread‑specific Random instance
                            Random rnd = new Random(Guid.NewGuid().GetHashCode());
                            for (int i = 0; i < readsPerThread; i++)
                            {
                                int row = rnd.Next(0, 1000);      // pick a random row
                                object val = readWs.Cells[row, 0].Value; // read cell value

                                // Expected value equals the row index
                                if (val == null || Convert.ToInt32(val) != row)
                                {
                                    Interlocked.Increment(ref inconsistencies);
                                }
                            }
                        }
                        catch
                        {
                            // Any unexpected exception is counted as an inconsistency
                            Interlocked.Increment(ref inconsistencies);
                        }
                        finally
                        {
                            // Signal completion of this thread
                            if (Interlocked.Increment(ref completedThreads) == threadCount)
                                done.Set();
                        }
                    });
                }

                // Wait for all threads to finish
                done.Wait();

                // -------------------------------------------------
                // Output results – any inconsistencies indicate data race issues
                // -------------------------------------------------
                Console.WriteLine($"Total reads performed : {threadCount * readsPerThread}");
                Console.WriteLine($"Inconsistencies detected: {inconsistencies}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
