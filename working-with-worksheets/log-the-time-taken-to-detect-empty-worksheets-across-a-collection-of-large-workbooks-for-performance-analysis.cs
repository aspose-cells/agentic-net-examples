// Title: Log performance of empty worksheet detection across large Excel workbooks with Aspose.Cells for .NET
// Description: A C# console app that iterates over a collection of Excel files, loads each workbook with Aspose.Cells, counts worksheets where Cells.MaxDataRow and Cells.MaxDataColumn are -1 (empty), and records per‑file and total elapsed time using Stopwatch. Ideal for benchmarking batch validation of large workbooks.
// Keywords: Aspose.Cells | C# | empty worksheet detection | performance measurement | Stopwatch | large Excel workbook | batch processing | Excel file validation | .NET benchmark | worksheet emptiness
// Common Searches: how to measure time for empty sheet detection with Aspose.Cells | benchmark Aspose.Cells loading large workbooks .NET | log per‑file processing time for Excel validation | detect and count empty worksheets in batch using C# | performance testing Aspose.Cells workbook parsing
// Developer Intent: Measure and log the elapsed time required to identify empty worksheets in each workbook and the aggregate time for the entire batch.
// Use Cases: Benchmarking Aspose.Cells load speed before optimizing data pipelines. | Skipping empty sheets in automated ETL jobs while tracking processing time. | Generating performance reports for enterprise Excel validation routines.
// AI Prompts: Provide C# code that records per‑workbook and overall detection time for empty worksheets using Aspose.Cells and prints a summary. | Suggest a thread‑safe way to process the workbook list in parallel and still capture accurate timing for each file. | Explain how Cells.MaxDataRow and Cells.MaxDataColumn can be used to reliably detect empty worksheets and log performance metrics.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    // A C# console app that iterates over a collection of Excel files, loads each workbook with Aspose.Cells, counts worksheets where Cells.MaxDataRow and Cells.MaxDataColumn are -1 (empty), and records per‑file and total elapsed time using Stopwatch. Ideal for benchmarking batch validation of large workbooks.
    class Program
    {
        static void Main()
        {
            // List of workbook file paths to be processed
            List<string> workbookFiles = new List<string>
            {
                @"C:\Data\LargeWorkbook1.xlsx",
                @"C:\Data\LargeWorkbook2.xlsx",
                @"C:\Data\LargeWorkbook3.xlsx"
                // Add more file paths as needed
            };

            // Stopwatch to measure total elapsed time
            Stopwatch totalStopwatch = Stopwatch.StartNew();

            foreach (string filePath in workbookFiles)
            {
                // Verify that the file exists before attempting to load it
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    Console.WriteLine();
                    continue;
                }

                // Measure time for each workbook
                Stopwatch wbStopwatch = Stopwatch.StartNew();

                try
                {
                    // Load the workbook (no custom interrupt monitor needed for this demo)
                    Workbook wb = new Workbook(filePath);

                    // Detect empty worksheets
                    int emptySheetCount = 0;
                    foreach (Worksheet sheet in wb.Worksheets)
                    {
                        // A worksheet is considered empty if it has no data rows and no data columns.
                        // Cells.MaxDataRow and Cells.MaxDataColumn return -1 when there is no data.
                        if (sheet.Cells.MaxDataRow < 0 && sheet.Cells.MaxDataColumn < 0)
                        {
                            emptySheetCount++;
                        }
                    }

                    wbStopwatch.Stop();
                    Console.WriteLine($"File: {filePath}");
                    Console.WriteLine($"  Empty worksheets: {emptySheetCount}");
                    Console.WriteLine($"  Detection time: {wbStopwatch.ElapsedMilliseconds} ms");
                    Console.WriteLine();
                }
                catch (Exception ex)
                {
                    wbStopwatch.Stop();
                    Console.WriteLine($"Error processing file: {filePath}");
                    Console.WriteLine($"  Exception: {ex.Message}");
                    Console.WriteLine();
                }
            }

            totalStopwatch.Stop();
            Console.WriteLine($"Total detection time for all workbooks: {totalStopwatch.ElapsedMilliseconds} ms");
        }
    }
}
