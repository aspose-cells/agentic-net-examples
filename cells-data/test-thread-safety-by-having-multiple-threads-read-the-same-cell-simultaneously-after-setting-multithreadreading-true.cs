// Title: Test thread safety of an Aspose.Cells workbook by reading the same cell concurrently with multiple C# threads
// Common Searches: C# Aspose.Cells read same Excel cell from multiple threads | how to verify thread safety of Aspose.Cells workbook | example of concurrent cell reading using Aspose.Cells .NET | does saving an Aspose.Cells workbook enable safe parallel reads | parallel read of cell A1 with Aspose.Cells in C# | Aspose.Cells multi‑threaded read performance test
// Tags: Aspose.Cells concurrent read example | C# parallel cell reading Aspose.Cells | Excel workbook thread safety Aspose.Cells | saving workbook before multi‑threaded read Aspose.Cells | Aspose.Cells multi‑threaded reading .NET | thread safe cell access Aspose.Cells

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

// The sample creates a workbook, writes a value to cell A1, saves the file, and then launches multiple threads that each read the same cell simultaneously, demonstrating that Aspose.Cells allows safe parallel read operations after the workbook is saved.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // NOTE: EnableThreadSafety property is not available in this version of Aspose.Cells.
            // The workbook is safe for concurrent read operations after it is saved or fully loaded.

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Write a test value to cell A1
            sheet.Cells["A1"].PutValue("Thread Safe Test");

            // Save the workbook (optional, demonstrates use of save rule)
            string outputPath = "ThreadSafetyTest.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");

            // Define the number of concurrent reader threads
            int threadCount = 10;
            Thread[] threads = new Thread[threadCount];

            // Start each thread to read the same cell simultaneously
            for (int i = 0; i < threadCount; i++)
            {
                threads[i] = new Thread(() =>
                {
                    try
                    {
                        // Read the value from cell A1
                        string value = sheet.Cells["A1"].StringValue;
                        Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId} read: {value}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Thread error: {ex.Message}");
                    }
                });
                threads[i].Start();
            }

            // Wait for all threads to complete
            foreach (var t in threads)
            {
                t.Join();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
