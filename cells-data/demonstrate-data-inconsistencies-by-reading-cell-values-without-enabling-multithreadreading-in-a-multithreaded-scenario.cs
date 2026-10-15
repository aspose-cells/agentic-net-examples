// Title: Demonstrate race‑condition data inconsistencies when reading Aspose.Cells worksheet cells concurrently without MultiThreadReading in C#
// AI Prompts: Write C# code that creates a workbook, fills a column with deterministic values, starts several threads that read the same cells without enabling workbook.Settings.MultiThreadReading, and logs any rows where the values differ between threads. | Show how to capture each thread's read results from an Aspose.Cells worksheet and compare them to identify race‑condition errors when MultiThreadReading is disabled.
// Common Searches: Aspose.Cells multithreaded read without MultiThreadReading example C# | detect inconsistent cell values when reading Excel with multiple threads .NET | C# race condition reading Aspose.Cells worksheet concurrently | why Aspose.Cells returns different values in parallel reads | how to test thread safety of Aspose.Cells cell access
// Tags: concurrent cell reading without MultiThreadReading Aspose.Cells | race condition detection in worksheet read C# | Aspose.Cells multi‑thread read safety test | cell value inconsistency detection .NET | disable MultiThreadReading Aspose.Cells example

using System;
using System.Threading;
using Aspose.Cells;

// The program creates a workbook, populates the first column with values 0‑999, launches multiple threads that read those cells without enabling MultiThreadReading, stores each thread's results, and then scans for any row where the values differ, reporting detected inconsistencies.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the first column with deterministic values (0..999)
        for (int i = 0; i < 1000; i++)
        {
            sheet.Cells[i, 0].PutValue(i);
        }

        // Number of concurrent reader threads
        int threadCount = 10;
        Thread[] threads = new Thread[threadCount];

        // Store read values per thread for later comparison
        int[,] readValues = new int[threadCount, 1000];

        // Launch threads WITHOUT enabling MultiThreadReading
        for (int t = 0; t < threadCount; t++)
        {
            int threadIndex = t; // capture loop variable
            threads[t] = new Thread(() =>
            {
                // Each thread reads all 1000 rows
                for (int row = 0; row < 1000; row++)
                {
                    // Directly read the cell value
                    object val = sheet.Cells[row, 0].Value;

                    // Convert to int (use -1 if null for safety)
                    readValues[threadIndex, row] = val == null ? -1 : Convert.ToInt32(val);
                }
            });
            threads[t].Start();
        }

        // Wait for all threads to finish
        foreach (Thread th in threads)
            th.Join();

        // Detect any inconsistencies among the threads' reads
        bool inconsistencyFound = false;
        for (int row = 0; row < 1000; row++)
        {
            int reference = readValues[0, row];
            for (int t = 1; t < threadCount; t++)
            {
                if (readValues[t, row] != reference)
                {
                    inconsistencyFound = true;
                    Console.WriteLine($"Inconsistency detected at row {row}: Thread0={reference}, Thread{t}={readValues[t, row]}");
                    break;
                }
            }
        }

        if (!inconsistencyFound)
        {
            Console.WriteLine("No inconsistency detected in this run (but reading without MultiThreadReading can cause race conditions).");
        }
    }
}
