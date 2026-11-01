// Title: Log and benchmark the time to detect empty worksheets in large Excel workbooks using Aspose.Cells for .NET
// Common Searches: how to measure execution time of empty worksheet detection with Aspose.Cells in C# | benchmarking Aspose.Cells performance for scanning large .xlsx files for empty sheets | C# code to log detection time of empty worksheets across multiple workbooks | performance analysis of worksheet emptiness check using Aspose.Cells .NET | record processing time per workbook when counting empty worksheets in Excel files
// Tags: Aspose.Cells empty worksheet detection timing | benchmarking worksheet emptiness with Aspose.Cells | performance measurement of workbook scanning .NET | large .xlsx file processing Aspose.Cells | stopwatch based timing of worksheet analysis

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPerformance
{
    // The sample loads each .xlsx file from a folder with Aspose.Cells, iterates through its worksheets, counts those where MaxDataRow is -1 (empty), and uses Stopwatch to log the detection time per workbook. Results—including empty worksheet count, detection time in milliseconds, and total processing time—are printed for performance analysis of large Excel workbooks.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the large workbooks to analyze
            string workbooksFolder = @"C:\LargeWorkbooks";

            // Verify that the directory exists
            if (!Directory.Exists(workbooksFolder))
            {
                Console.WriteLine($"Folder not found: {workbooksFolder}");
                return;
            }

            // Get all Excel files in the directory (you can adjust the search pattern as needed)
            string[] workbookFiles = Directory.GetFiles(workbooksFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            // List to hold timing results for each workbook
            var results = new List<WorkbookResult>();

            // Overall stopwatch for the entire operation
            Stopwatch totalStopwatch = Stopwatch.StartNew();

            foreach (string filePath in workbookFiles)
            {
                // Ensure the file still exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found, skipping: {filePath}");
                    continue;
                }

                try
                {
                    // Stopwatch for the current workbook
                    Stopwatch sw = Stopwatch.StartNew();

                    // Load the workbook (Aspose.Cells handles large files efficiently)
                    Workbook workbook = new Workbook(filePath);

                    int emptyWorksheetCount = 0;

                    // Iterate through all worksheets in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // A worksheet is considered empty if it has no data cells.
                        // MaxDataRow returns -1 when there are no data rows.
                        if (sheet.Cells.MaxDataRow == -1)
                        {
                            emptyWorksheetCount++;
                        }
                    }

                    sw.Stop();

                    // Store the result for this workbook
                    results.Add(new WorkbookResult
                    {
                        FileName = Path.GetFileName(filePath),
                        EmptyWorksheetCount = emptyWorksheetCount,
                        DetectionTimeMs = sw.ElapsedMilliseconds
                    });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            totalStopwatch.Stop();

            // Output the performance analysis
            Console.WriteLine("Performance Analysis of Empty Worksheet Detection");
            Console.WriteLine("-------------------------------------------------");
            foreach (var r in results)
            {
                Console.WriteLine($"Workbook: {r.FileName}");
                Console.WriteLine($"  Empty Worksheets : {r.EmptyWorksheetCount}");
                Console.WriteLine($"  Detection Time   : {r.DetectionTimeMs} ms");
                Console.WriteLine();
            }

            Console.WriteLine($"Total time for processing {workbookFiles.Length} workbooks: {totalStopwatch.ElapsedMilliseconds} ms");
        }

        // Helper class to hold per‑workbook results
        class WorkbookResult
        {
            public string FileName { get; set; }
            public int EmptyWorksheetCount { get; set; }
            public long DetectionTimeMs { get; set; }
        }
    }
}
