// Title: Benchmark cell iteration speed with and without MaxDataRow/MaxDataColumn limits in Aspose.Cells for C#
// AI Prompts: Write a C# console app that populates an Aspose.Cells worksheet with 5,000 rows and 100 columns of sample data, then uses Stopwatch to time a full‑range cell loop versus a loop bounded by Cells.MaxDataRow and Cells.MaxDataColumn. | Add try‑catch error handling to the benchmark program and output the elapsed milliseconds for both the unrestricted and limited iterations, clearly showing the time saved. | Extend the code to print the detected MaxDataRow and MaxDataColumn values, compare them to the total row and column counts, and then perform the performance measurement.
// Common Searches: how to compare full worksheet iteration time to MaxDataRow limited loop using Aspose.Cells in C# | Aspose.Cells performance test for cell enumeration with MaxDataRow and MaxDataColumn | C# measure speed difference between Cells.Rows.Count loop and Cells.MaxDataRow loop | benchmarking Aspose.Cells cell traversal using Stopwatch in .NET
// Tags: benchmark Aspose.Cells cell iteration performance | use MaxDataRow MaxDataColumn for efficient worksheet traversal | measure iteration time with Stopwatch in Aspose.Cells | optimize Excel data enumeration using data bounds in .NET | performance comparison full range vs bounded cell loops

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example creates a Workbook, fills a 5,000‑row by 100‑column worksheet with test values, recalculates formulas, and then measures two nested loops: one iterating over the full Rows.Count and Columns.Count range, and another iterating only up to Cells.MaxDataRow and Cells.MaxDataColumn. Stopwatch records the elapsed time for each approach, and the program outputs both timings and the time saved, with exception handling for robustness.
class MaxDataLimitsPerformanceTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define size of test data
            const int totalRows = 5000;
            const int totalCols = 100;

            // Fill the worksheet with sample data
            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    sheet.Cells[r, c].PutValue($"R{r}C{c}");
                }
            }

            // Recalculate formulas (if any)
            workbook.CalculateFormula();

            // -----------------------------------------------------------------
            // Measure iteration time without using MaxDataRow/MaxDataColumn limits
            // (iterate over the full possible range of rows and columns)
            // -----------------------------------------------------------------
            Stopwatch swFull = Stopwatch.StartNew();

            // Use the total rows/columns count from the Cells collection
            int fullRowCount = sheet.Cells.Rows.Count;
            int fullColCount = sheet.Cells.Columns.Count;

            for (int r = 0; r < fullRowCount; r++)
            {
                for (int c = 0; c < fullColCount; c++)
                {
                    // Access the cell value (no processing needed for timing)
                    var val = sheet.Cells[r, c].Value;
                }
            }

            swFull.Stop();
            Console.WriteLine($"Full iteration time (rows={fullRowCount}, cols={fullColCount}): {swFull.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Measure iteration time using MaxDataRow and MaxDataColumn limits
            // (iterate only over the region that actually contains data)
            // -----------------------------------------------------------------
            Stopwatch swLimited = Stopwatch.StartNew();

            // MaxDataRow and MaxDataColumn are zero‑based indexes of the last cell that contains data
            int maxDataRow = sheet.Cells.MaxDataRow;       // e.g., totalRows - 1
            int maxDataCol = sheet.Cells.MaxDataColumn;   // e.g., totalCols - 1

            for (int r = 0; r <= maxDataRow; r++)
            {
                for (int c = 0; c <= maxDataCol; c++)
                {
                    var val = sheet.Cells[r, c].Value;
                }
            }

            swLimited.Stop();
            Console.WriteLine($"Limited iteration time (maxRow={maxDataRow + 1}, maxCol={maxDataCol + 1}): {swLimited.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Output the difference
            // -----------------------------------------------------------------
            long diff = swFull.ElapsedMilliseconds - swLimited.ElapsedMilliseconds;
            Console.WriteLine($"Time saved by using MaxData limits: {diff} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
