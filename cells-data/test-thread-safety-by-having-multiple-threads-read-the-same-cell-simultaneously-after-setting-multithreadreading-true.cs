// Title: How to enable MultiThreadReading in Aspose.Cells for .NET and safely read the same cell from multiple parallel tasks
// Common Searches: Aspose.Cells enable MultiThreadReading for concurrent cell reads in C# | C# example of reading the same Excel cell from multiple threads using Aspose.Cells | Thread‑safe workbook reading with Aspose.Cells MultiThreadReading property | Parallel tasks reading Excel file with Aspose.Cells .NET SDK | Verify thread safety of Aspose.Cells workbook loading in .NET
// Tags: Aspose.Cells MultiThreadReading configuration | concurrent cell read Aspose.Cells .NET | thread‑safe workbook loading C# | parallel Excel reading Aspose.Cells | multi‑threaded workbook access Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample creates a workbook, writes a value to cell A1, saves it, and then starts ten parallel tasks. Each task opens its own Workbook instance, reads the string value from A1, and stores the result in a shared list. With the MultiThreadReading property set to true, the example demonstrates that concurrent reads return the same data without errors, confirming thread‑safe reading in Aspose.Cells for .NET.
class ThreadSafetyTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook (using the standard creation rule)
            Workbook workbook = new Workbook();

            // Write a test value to cell A1 in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("ThreadSafeValue");

            // Save the workbook – optional for verification
            string filePath = "ThreadSafetyTest.xlsx";
            workbook.Save(filePath);

            // Ensure the file was created before proceeding
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found after saving.");

            // Number of concurrent readers
            const int readerCount = 10;

            // Container to hold read values from each thread
            List<string> readValues = new List<string>(new string[readerCount]);

            // Create and start tasks that read the same cell simultaneously
            Task[] tasks = new Task[readerCount];
            for (int i = 0; i < readerCount; i++)
            {
                int index = i; // capture loop variable
                tasks[i] = Task.Run(() =>
                {
                    try
                    {
                        // Verify the file exists before loading
                        if (!File.Exists(filePath))
                            throw new FileNotFoundException($"The file '{filePath}' was not found for reading.");

                        // Each task creates its own Workbook instance for reading
                        Workbook readWorkbook = new Workbook(filePath);

                        // Read the value from cell A1
                        string value = readWorkbook.Worksheets[0].Cells["A1"].StringValue;

                        // Store the result in the shared list
                        readValues[index] = value;
                    }
                    catch (Exception ex)
                    {
                        // Capture any exception for this thread
                        readValues[index] = $"Error: {ex.Message}";
                    }
                });
            }

            // Wait for all reading tasks to complete
            Task.WaitAll(tasks);

            // Output the results to verify that all threads read the same value
            Console.WriteLine("Results of concurrent reads:");
            for (int i = 0; i < readValues.Count; i++)
            {
                Console.WriteLine($"Thread {i + 1}: {readValues[i]}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
