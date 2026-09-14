// Title: Compare cell enumeration speed using Cells.MaxDisplayRange versus full Cells.GetEnumerator() in a large Aspose.Cells worksheet (C#)
// AI Prompts: Execute the sample program and log the elapsed milliseconds for enumerating cells with MaxDisplayRange and with the full Cells enumerator, then display the speed difference. | Update the benchmark to run each enumeration loop 10 times, calculate the average duration for MaxDisplayRange and full sheet traversal, and print the aggregated timings. | Instrument the code to record peak memory usage during each enumeration using System.Diagnostics.Process and include memory consumption in the benchmark output.
// Common Searches: Aspose.Cells C# benchmark MaxDisplayRange enumeration versus full sheet iteration | how fast is Cells.MaxDisplayRange.GetEnumerator() compared to Cells.GetEnumerator() in .NET | measure cell iteration performance in large worksheet using Aspose.Cells | performance testing cell enumeration with Aspose.Cells MaxDisplayRange
// Tags: Cells.MaxDisplayRange enumeration speed | full sheet Cells.GetEnumerator performance | large worksheet cell iteration benchmark | Aspose.Cells enumeration timing C# | stopwatch performance measurement Aspose.Cells

using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsBenchmark
{
    // The example creates a 5,000 × 100 worksheet, fills it with numeric data, then uses Stopwatch to time cell enumeration via Cells.MaxDisplayRange.GetEnumerator() and via Cells.GetEnumerator(), prints both elapsed times, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate a large dataset (e.g., 5000 rows x 100 columns)
                const int totalRows = 5000;
                const int totalCols = 100;
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = 0; col < totalCols; col++)
                    {
                        cells[row, col].PutValue(row * totalCols + col);
                    }
                }

                // Ensure the MaxDisplayRange is calculated (it includes data, merged cells, shapes)
                AsposeRange maxDisplayRange = cells.MaxDisplayRange;

                // Benchmark enumeration using MaxDisplayRange
                Stopwatch sw = new Stopwatch();
                sw.Start();
                IEnumerator rangeEnum = maxDisplayRange.GetEnumerator();
                while (rangeEnum.MoveNext())
                {
                    // Access the cell (cast to Cell) – no operation needed, just enumeration
                    Cell cell = (Cell)rangeEnum.Current;
                    // Example: read the value (optional, keeps the loop realistic)
                    var val = cell.Value;
                }
                sw.Stop();
                TimeSpan timeUsingDisplayRange = sw.Elapsed;

                // Benchmark full sheet traversal using Cells.GetEnumerator()
                sw.Restart();
                IEnumerator sheetEnum = cells.GetEnumerator();
                while (sheetEnum.MoveNext())
                {
                    Cell cell = (Cell)sheetEnum.Current;
                    var val = cell.Value;
                }
                sw.Stop();
                TimeSpan timeFullSheet = sw.Elapsed;

                // Output the benchmark results
                Console.WriteLine($"Enumeration using MaxDisplayRange: {timeUsingDisplayRange.TotalMilliseconds} ms");
                Console.WriteLine($"Enumeration using full sheet traversal: {timeFullSheet.TotalMilliseconds} ms");

                // Save the workbook (optional, demonstrates that data persists)
                string outputPath = "BenchmarkResult.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during execution: {ex.Message}");
            }
        }
    }
}
