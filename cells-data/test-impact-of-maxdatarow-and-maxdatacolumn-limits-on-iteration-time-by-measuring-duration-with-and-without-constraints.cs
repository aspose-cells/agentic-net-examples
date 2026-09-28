// Title: Measure Aspose.Cells worksheet iteration performance with and without MaxDataRow/MaxDataColumn limits in C#
// AI Prompts: Write a C# program that fills an Aspose.Cells worksheet with 5,000 rows and 200 columns, then uses Stopwatch to time a nested loop that iterates only up to Cells.MaxDataRow and Cells.MaxDataColumn. | Add a second timing loop that iterates over the full 5,000 × 200 range, compare the elapsed milliseconds with the limited loop, and print both results.
// Common Searches: Aspose.Cells C# how long does iterating cells up to MaxDataRow take compared to full range | benchmark worksheet cell traversal using MaxDataColumn limit in Aspose.Cells | performance test for Aspose.Cells iteration with large dataset 5000 rows 200 columns | measure impact of MaxDataRow and MaxDataColumn on loop execution time in C# | compare limited vs full range cell iteration speed Aspose.Cells
// Tags: Aspose.Cells iteration performance measurement | C# benchmark MaxDataRow MaxDataColumn | cell traversal timing Aspose.Cells | limited range vs full range worksheet iteration | performance testing large worksheet Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace MaxDataIterationTest
{
    // The example creates a workbook, populates it with 5,000 rows and 200 columns of sample data, then uses Stopwatch to record the time of two nested loops: one bounded by Cells.MaxDataRow/MaxDataColumn and another that iterates over the entire defined range. It prints both elapsed times and saves the workbook, demonstrating how the MaxDataRow/MaxDataColumn limits affect iteration speed.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define a relatively large data set
            const int totalRows = 5000;
            const int totalCols = 200;

            // Populate the worksheet with sample data
            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    cells[r, c].PutValue(r * totalCols + c);
                }
            }

            // Retrieve the MaxDataRow and MaxDataColumn after data insertion
            int maxDataRow = cells.MaxDataRow;       // zero‑based index of the last row containing data
            int maxDataColumn = cells.MaxDataColumn; // zero‑based index of the last column containing data

            // -----------------------------------------------------------------
            // Measure iteration time using MaxDataRow / MaxDataColumn limits
            // -----------------------------------------------------------------
            Stopwatch swLimited = new Stopwatch();
            swLimited.Start();

            for (int r = 0; r <= maxDataRow; r++)
            {
                for (int c = 0; c <= maxDataColumn; c++)
                {
                    // Access the cell value (no processing needed for timing)
                    var _ = cells[r, c].Value;
                }
            }

            swLimited.Stop();
            Console.WriteLine($"Iteration with limits (rows: {maxDataRow + 1}, cols: {maxDataColumn + 1}) took {swLimited.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Measure iteration time without using the limits (full range)
            // -----------------------------------------------------------------
            Stopwatch swFull = new Stopwatch();
            swFull.Start();

            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    var _ = cells[r, c].Value;
                }
            }

            swFull.Stop();
            Console.WriteLine($"Iteration without limits (rows: {totalRows}, cols: {totalCols}) took {swFull.ElapsedMilliseconds} ms");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("MaxDataIterationDemo.xlsx");
        }
    }
}
