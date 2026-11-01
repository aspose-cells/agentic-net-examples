// Title: Measure enumeration performance of Aspose.Cells MaxDisplayRange versus full worksheet traversal in C#
// AI Prompts: Execute the sample program and capture the elapsed milliseconds for the MaxDisplayRange loop, then contrast it with the nested for‑loop over MaxDataRow/MaxDataColumn. | Extend the benchmark to accept rowCount and colCount as command‑line parameters, generate a workbook of the specified size, and report the percentage speed difference between used‑range iteration and full sheet traversal. | Instrument the code to calculate and log the average time per cell read for each enumeration method, outputting the results to a CSV file for further analysis.
// Common Searches: aspnet benchmark Aspose.Cells used range iteration speed | how fast is MaxDisplayRange enumeration compared to full sheet loop in Aspose.Cells | C# performance test reading cells with MaxDataRow and MaxDataColumn versus MaxDisplayRange | measure cell read latency in large Aspose.Cells workbook | Aspose.Cells performance comparison of DisplayRange and full worksheet traversal
// Tags: Aspose.Cells MaxDisplayRange iteration speed | full worksheet traversal latency Aspose.Cells | cell read benchmark .NET | used range performance testing Aspose.Cells | C# workbook enumeration performance

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsBenchmark
{
    // The example creates a 5,000‑row by 100‑column workbook, fills each cell with an integer, then benchmarks two read strategies: iterating over Worksheet.Cells.MaxDisplayRange and looping through all rows and columns up to MaxDataRow/MaxDataColumn. It prints the elapsed time for each method, allowing developers to evaluate the performance gain of using the used‑range enumeration.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the sheet with sample data (e.g., 5000 rows x 100 columns)
                const int totalRows = 5000;
                const int totalCols = 100;
                for (int r = 0; r < totalRows; r++)
                {
                    for (int c = 0; c < totalCols; c++)
                    {
                        // Fill each cell with a simple integer value
                        sheet.Cells[r, c].PutValue(r * totalCols + c);
                    }
                }

                // Adjust column widths and row heights for the used range
                sheet.AutoFitColumns();
                sheet.AutoFitRows();

                // Benchmark iterating over the DisplayRange (used range)
                Stopwatch swDisplay = Stopwatch.StartNew();
                Aspose.Cells.Range displayRange = sheet.Cells.MaxDisplayRange; // Equivalent to DisplayRange
                foreach (Cell cell in displayRange)
                {
                    // Simple read operation (no modification)
                    var val = cell.Value;
                }
                swDisplay.Stop();

                // Benchmark iterating over the entire sheet (full traversal)
                Stopwatch swFull = Stopwatch.StartNew();
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;
                for (int r = 0; r <= maxRow; r++)
                {
                    for (int c = 0; c <= maxCol; c++)
                    {
                        var val = sheet.Cells[r, c].Value;
                    }
                }
                swFull.Stop();

                // Output the timing results
                Console.WriteLine($"DisplayRange iteration time: {swDisplay.ElapsedMilliseconds} ms");
                Console.WriteLine($"Full sheet traversal time: {swFull.ElapsedMilliseconds} ms");

                // Optional: Save the workbook to verify data (not part of benchmark)
                workbook.Save("BenchmarkResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
