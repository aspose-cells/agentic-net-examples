// Title: Benchmark Cells.MaxDisplayRange vs full sheet enumeration in Aspose.Cells for .NET
// Description: Creates a 10,000 × 50 workbook, fills each cell, forces MaxDisplayRange calculation, then measures and compares the elapsed time of iterating cells with MaxDisplayRange.GetEnumerator() and Cells.GetEnumerator(). The results are printed and the workbook can be saved.
// Keywords: Aspose.Cells | C# | .NET | benchmark | enumeration performance | MaxDisplayRange | Cells.GetEnumerator | spreadsheet traversal speed | large worksheet processing | cell iteration timing
// Common Searches: Aspose.Cells benchmark MaxDisplayRange enumeration speed | compare Cells.GetEnumerator vs MaxDisplayRange performance | measure cell traversal time in large Excel files using Aspose.Cells | how fast is MaxDisplayRange enumeration in C# | optimize spreadsheet iteration with Aspose.Cells
// Developer Intent: Evaluate whether iterating through Cells.MaxDisplayRange provides a measurable speed advantage over iterating the entire worksheet with Cells.GetEnumerator.
// Use Cases: Choose the most efficient enumeration method for data‑intensive Excel processing. | Generate performance baselines before implementing bulk cell operations. | Validate optimization decisions in ETL pipelines that manipulate large spreadsheets.
// AI Prompts: Write a reusable C# function that accepts row and column counts, runs the MaxDisplayRange vs full sheet enumeration benchmark, and returns the time difference. | Suggest ways to reduce enumeration overhead when processing millions of cells with Aspose.Cells. | Explain how MaxDisplayRange is determined and why it can be faster than a full sheet enumeration.

using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsBenchmark
{
    // Creates a 10,000 × 50 workbook, fills each cell, forces MaxDisplayRange calculation, then measures and compares the elapsed time of iterating cells with MaxDisplayRange.GetEnumerator() and Cells.GetEnumerator(). The results are printed and the workbook can be saved.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate a large dataset (e.g., 10,000 rows x 50 columns)
                const int totalRows = 10000;
                const int totalCols = 50;
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = 0; col < totalCols; col++)
                    {
                        cells[row, col].PutValue(row * totalCols + col);
                    }
                }

                // Ensure the MaxDisplayRange is calculated
                AsposeRange maxDisplayRange = cells.MaxDisplayRange;

                // Benchmark enumeration using MaxDisplayRange
                Stopwatch sw = new Stopwatch();
                sw.Start();
                IEnumerator displayEnum = maxDisplayRange.GetEnumerator();
                while (displayEnum.MoveNext())
                {
                    Cell cell = (Cell)displayEnum.Current;
                    // Access the cell value to simulate realistic work
                    var _ = cell.Value;
                }
                sw.Stop();
                long displayRangeTime = sw.ElapsedMilliseconds;

                // Benchmark enumeration over the entire sheet
                sw.Restart();
                IEnumerator fullEnum = cells.GetEnumerator();
                while (fullEnum.MoveNext())
                {
                    Cell cell = (Cell)fullEnum.Current;
                    var _ = cell.Value;
                }
                sw.Stop();
                long fullSheetTime = sw.ElapsedMilliseconds;

                // Output the benchmark results
                Console.WriteLine($"Enumeration using MaxDisplayRange: {displayRangeTime} ms");
                Console.WriteLine($"Enumeration over full sheet: {fullSheetTime} ms");

                // Save the workbook (optional, just to keep the data)
                string outputPath = "BenchmarkResult.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during execution: {ex.Message}");
            }
        }
    }
}
