// Title: Aspose.Cells C# Demo: Data Inconsistencies When Reading Cells Concurrently Without MultiThreadReading
// Description: This example creates a workbook, fills column A with sequential integers, and starts five threads that read distinct row ranges without enabling MultiThreadReading. Each thread checks that the retrieved value matches the expected row index and logs any mismatches or exceptions, illustrating the race conditions that can occur when concurrent reads are performed without the MultiThreadReading flag.
// Keywords: Aspose.Cells | C# multithread reading | MultiThreadReading disabled | cell read race condition | thread safety Aspose.Cells | concurrent cell access | data inconsistency example | Aspose.Cells performance testing
// Common Searches: Aspose.Cells read cells from multiple threads without MultiThreadReading | Why does Aspose.Cells show inconsistent values in multithreaded reads | Example of race condition in Aspose.Cells when MultiThreadReading is false | How to reproduce data mismatch in Aspose.Cells multithreading | Aspose.Cells thread safety demo C#
// Developer Intent: Demonstrate that reading cell values from several threads without enabling MultiThreadReading can lead to inconsistent results.
// Use Cases: Validate thread‑safety of cell reads before enabling MultiThreadReading in high‑throughput spreadsheet pipelines. | Create a benchmark that compares correctness and speed with MultiThreadReading disabled versus enabled. | Generate logs of mismatched values to diagnose race conditions in multithreaded Aspose.Cells applications.
// AI Prompts: Rewrite the sample to enable MultiThreadReading and show that no inconsistencies are reported. | Produce an NUnit test that asserts zero mismatches when MultiThreadReading is set to true for the same multithreaded scenario. | Explain the internal synchronization mechanism Aspose.Cells uses when MultiThreadReading is enabled and why it is required for concurrent reads.

using System;
using System.Text;
using System.Threading;
using Aspose.Cells;

// This example creates a workbook, fills column A with sequential integers, and starts five threads that read distinct row ranges without enabling MultiThreadReading. Each thread checks that the retrieved value matches the expected row index and logs any mismatches or exceptions, illustrating the race conditions that can occur when concurrent reads are performed without the MultiThreadReading flag.
class MultiThreadReadingDemo
{
    static void Main()
    {
        // Create a new workbook and get the cells collection
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Populate the first column with sequential integer values
        int totalRows = 1000;
        for (int i = 0; i < totalRows; i++)
        {
            cells[i, 0].PutValue(i);
        }

        // NOTE: MultiThreadReading is NOT enabled (default is false)

        int threadCount = 5;
        int rowsPerThread = totalRows / threadCount;
        int finishedThreads = 0;
        StringBuilder errorLog = new StringBuilder();

        // Launch multiple threads that read cell values concurrently
        for (int t = 0; t < threadCount; t++)
        {
            int startRow = t * rowsPerThread;
            int endRow = (t == threadCount - 1) ? totalRows : startRow + rowsPerThread;

            Thread thread = new Thread(() =>
            {
                try
                {
                    for (int r = startRow; r < endRow; r++)
                    {
                        // Read the cell value
                        object val = cells[r, 0].Value;

                        // Verify that the value matches the expected row index
                        if (val == null || Convert.ToInt32(val) != r)
                        {
                            lock (errorLog)
                            {
                                errorLog.AppendLine($"Inconsistent value at row {r}: expected {r}, got {val ?? "null"}");
                            }
                        }
                    }
                    Interlocked.Increment(ref finishedThreads);
                }
                catch (Exception ex)
                {
                    lock (errorLog)
                    {
                        errorLog.AppendLine($"Thread {Thread.CurrentThread.ManagedThreadId} exception: {ex.Message}");
                    }
                }
            });

            thread.Start();
        }

        // Wait for all threads to complete
        while (Interlocked.CompareExchange(ref finishedThreads, 0, 0) < threadCount)
        {
            Thread.Sleep(100);
        }

        // Output any detected inconsistencies
        if (errorLog.Length > 0)
        {
            Console.WriteLine("Data inconsistencies detected:");
            Console.WriteLine(errorLog.ToString());
        }
        else
        {
            Console.WriteLine("All threads read consistent data (unexpected when MultiThreadReading is disabled).");
        }
    }
}
