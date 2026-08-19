// Title: C# Parallel Smart Marker Processing with Aspose.Cells for Large Datasets
// Description: Loads a template workbook, creates 200,000 Person objects, assigns them to a WorkbookDesigner, splits the smart‑marker range into sub‑ranges based on CPU cores, processes each sub‑range concurrently using separate threads, and saves the populated Excel file.
// Keywords: Aspose.Cells | smart markers | parallel processing | multithreaded Excel generation | C# | large data set | WorkbookDesigner | high‑performance reporting
// Common Searches: Aspose.Cells parallel smart marker example | multithreaded smart marker processing C# | populate 200k rows with smart markers | speed up Excel generation Aspose.Cells | threaded workbook designer usage
// Developer Intent: Speed up smart‑marker expansion by dividing the worksheet into independent ranges and processing them on multiple threads.
// Use Cases: Generate an employee directory with hundreds of thousands of rows in minutes. | Produce high‑volume financial statements where each sheet segment is filled concurrently. | Implement a web service that receives large collections and returns Excel reports with minimal latency.
// AI Prompts: Show how to rewrite the example using Task Parallel Library (TPL) instead of raw Thread objects. | Explain thread‑safety considerations when sharing a WorkbookDesigner across parallel tasks. | Suggest memory‑optimisation techniques for processing massive smart‑marker ranges in parallel.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Cells;

// Alias to avoid ambiguity between Aspose.Cells.Range and System.Range.
using AsposeRange = Aspose.Cells.Range;

// Loads a template workbook, creates 200,000 Person objects, assigns them to a WorkbookDesigner, splits the smart‑marker range into sub‑ranges based on CPU cores, processes each sub‑range concurrently using separate threads, and saves the populated Excel file.
class ParallelSmartMarkerProcessing
{
    static void Main()
    {
        try
        {
            const string templatePath = "TemplateWithSmartMarkers.xlsx";
            const string resultPath = "ResultParallel.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException.
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load a template workbook that contains smart markers.
            Workbook workbook = new Workbook(templatePath);
            Worksheet sheet = workbook.Worksheets[0];

            // Generate a massive data set.
            List<Person> persons = GeneratePersons(200_000);

            // Set the data source for the smart markers.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Persons", persons);

            // Define a range that covers the area where the smart markers will be expanded.
            // Assume the template has markers starting from row 2 (index 1) in columns A‑C.
            int totalRows = persons.Count + 1; // +1 for the header row in the template
            AsposeRange fullRange = sheet.Cells.CreateRange(1, 0, totalRows, 3);
            fullRange.Name = "_CellsSmartMarkers";

            // Determine how many threads to run (use the number of logical processors).
            int threadCount = Environment.ProcessorCount;
            int rowsPerThread = Math.Max(1, totalRows / threadCount);
            Thread[] threads = new Thread[threadCount];

            for (int i = 0; i < threadCount; i++)
            {
                // Calculate the start row for this thread (zero‑based index).
                int startRow = 1 + i * rowsPerThread;
                // The last thread may take the remaining rows.
                int endRow = (i == threadCount - 1) ? (1 + totalRows) : startRow + rowsPerThread;

                // Create a sub‑range for the current thread.
                AsposeRange subRange = sheet.Cells.CreateRange(startRow, 0, endRow - startRow, 3);
                subRange.Name = $"_CellsSmartMarkers_{i}";

                // Launch the thread to process only its sub‑range.
                threads[i] = new Thread(() =>
                {
                    try
                    {
                        // The boolean parameter indicates whether to preserve unrecognized markers.
                        designer.Process(subRange, true);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Thread error: {ex.Message}");
                    }
                });
                threads[i].Start();
            }

            // Wait for all threads to complete.
            foreach (Thread t in threads) t.Join();

            // Save the populated workbook.
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to create a large list of sample data.
    static List<Person> GeneratePersons(int count)
    {
        var list = new List<Person>(count);
        for (int i = 0; i < count; i++)
        {
            list.Add(new Person
            {
                Name = $"Person {i}",
                Age = 20 + (i % 30),
                Department = $"Dept {(i % 5) + 1}"
            });
        }
        return list;
    }

    // Simple POCO used as a data source for smart markers.
    public class Person
    {
        public string? Name { get; set; }
        public int Age { get; set; }
        public string? Department { get; set; }
    }
}
