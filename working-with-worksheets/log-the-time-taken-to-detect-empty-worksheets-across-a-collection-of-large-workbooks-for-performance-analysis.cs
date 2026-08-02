// Title: Measure empty worksheet detection time in large Excel workbooks with Aspose.Cells (C#)
// Description: C# program that iterates a collection of Excel files, loads each workbook with Aspose.Cells, counts worksheets with no data (MaxDataRow and MaxDataColumn = -1), records the elapsed milliseconds using Stopwatch, and logs the file name, empty‑sheet count, and detection time to the console while handling missing files and exceptions.
// Keywords: Aspose.Cells empty worksheet detection | C# benchmark Excel sheet analysis | measure worksheet detection time | stopwatch performance Aspose.Cells | large workbook processing .NET | count empty sheets Aspose.Cells | Excel performance logging C#
// Common Searches: how to time empty worksheet detection with Aspose.Cells | benchmarking Aspose.Cells worksheet analysis | C# measure performance of empty sheet count | log detection time for empty sheets in large Excel files | Aspose.Cells performance tips for batch workbook processing
// Developer Intent: Record how long it takes to identify empty worksheets in each large workbook.
// Use Cases: Profile and optimize batch validation of Excel files by measuring empty‑sheet detection latency. | Generate per‑file performance reports for data‑import pipelines that need to skip blank worksheets. | Integrate timing metrics into monitoring dashboards to detect regressions in workbook processing speed.
// AI Prompts: Rewrite the example to output results to a CSV file with columns: FilePath, EmptyWorksheetCount, DetectionMs. | Show how to process the workbook list in parallel while preserving accurate timing for each file. | Suggest memory‑efficient loading strategies for very large workbooks that still allow empty‑sheet detection and timing.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// C# program that iterates a collection of Excel files, loads each workbook with Aspose.Cells, counts worksheets with no data (MaxDataRow and MaxDataColumn = -1), records the elapsed milliseconds using Stopwatch, and logs the file name, empty‑sheet count, and detection time to the console while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        // List of workbook file paths to be processed
        var workbookFiles = new List<string>
        {
            "LargeWorkbook1.xlsx",
            "LargeWorkbook2.xlsx",
            "LargeWorkbook3.xlsx"
        };

        foreach (var filePath in workbookFiles)
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                Console.WriteLine(new string('-', 40));
                continue;
            }

            try
            {
                // Start measuring time for the detection operation
                Stopwatch timer = Stopwatch.StartNew();

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Count empty worksheets
                int emptyWorksheetCount = 0;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // A worksheet is considered empty when it contains no data cells.
                    // MaxDataRow and MaxDataColumn return -1 when there is no data.
                    if (sheet.Cells.MaxDataRow < 0 && sheet.Cells.MaxDataColumn < 0)
                    {
                        emptyWorksheetCount++;
                    }
                }

                // Stop timing
                timer.Stop();

                // Log the result
                Console.WriteLine($"File: {filePath}");
                Console.WriteLine($"Empty worksheets: {emptyWorksheetCount}");
                Console.WriteLine($"Detection time: {timer.ElapsedMilliseconds} ms");
                Console.WriteLine(new string('-', 40));
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors during processing
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                Console.WriteLine(new string('-', 40));
            }
        }
    }
}
