// Title: Measure performance of empty worksheet detection across multiple large .xlsx workbooks using Aspose.Cells in C#
// AI Prompts: Create a C# console application that walks through a directory of .xlsx files, loads each workbook with Aspose.Cells, counts worksheets where MaxDataRow and MaxDataColumn are -1, and records the total elapsed time for the scan. | Update the worksheet‑iteration loop to capture the processing duration for each individual workbook and display it together with the empty worksheet count. | Add code to generate a CSV report that lists the workbook name, number of empty worksheets, and the processing time for each file.
// Common Searches: how long does Aspose.Cells take to scan large Excel files for empty sheets in C# | C# code to benchmark empty worksheet detection across many .xlsx workbooks | measure performance of worksheet scanning with Aspose.Cells .NET | log processing time while counting empty worksheets in multiple Excel workbooks | Aspose.Cells performance test for empty sheet detection in a folder of workbooks
// Tags: Aspose.Cells empty sheet detection benchmark | C# worksheet scanning performance measurement | large .xlsx workbook processing with Aspose.Cells | measure Excel sheet analysis time in .NET | log workbook scan duration using Aspose.Cells

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The program scans all .xlsx files in a specified folder, counts worksheets that contain no data using Aspose.Cells, logs total and per‑workbook processing times, and can export the results to a CSV file.
class EmptyWorksheetPerformanceLogger
{
    static void Main()
    {
        try
        {
            // Directory containing the large workbooks to analyze
            string workbooksFolder = @"C:\LargeWorkbooks";

            // Verify that the directory exists
            if (!Directory.Exists(workbooksFolder))
            {
                Console.WriteLine($"Directory not found: {workbooksFolder}");
                return;
            }

            // Gather all workbook file paths (e.g., .xlsx files)
            List<string> workbookFiles = new List<string>(Directory.GetFiles(workbooksFolder, "*.xlsx", SearchOption.AllDirectories));

            // Stopwatch to measure total detection time
            Stopwatch totalStopwatch = Stopwatch.StartNew();

            // Store results for each workbook
            var results = new List<(string FileName, int EmptyWorksheetCount)>();

            foreach (string filePath in workbookFiles)
            {
                // Ensure the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found, skipping: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (Aspose.Cells handles large files efficiently)
                    Workbook workbook = new Workbook(filePath);

                    int emptyCount = 0;

                    // Iterate through all worksheets in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // A worksheet is considered empty when it has no data rows and no data columns
                        // MaxDataRow and MaxDataColumn return -1 if the sheet contains no data
                        if (sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1)
                        {
                            emptyCount++;
                        }
                    }

                    results.Add((Path.GetFileName(filePath), emptyCount));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            totalStopwatch.Stop();

            // Output performance summary
            Console.WriteLine("=== Empty Worksheet Detection Performance ===");
            Console.WriteLine($"Total workbooks processed: {results.Count}");
            Console.WriteLine($"Total time elapsed: {totalStopwatch.Elapsed.TotalSeconds:F2} seconds");
            Console.WriteLine();

            // Detailed per‑workbook results
            foreach (var result in results)
            {
                Console.WriteLine($"Workbook: {result.FileName} - Empty worksheets: {result.EmptyWorksheetCount}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
