// Title: C# Benchmark: Iterating Cells with MaxDataRow/MaxDataColumn vs Fixed Oversized Range in Aspose.Cells
// Description: This example creates a workbook with 5,000 rows and 100 columns, saves it, and measures the time required to read all populated cells using the dynamic MaxDataRow/MaxDataColumn limits. It then repeats the read operation over a deliberately larger range (10,000 × 200) using CheckCell, allowing a direct performance comparison between constrained and unconstrained iteration.
// Keywords: Aspose.Cells | MaxDataRow | MaxDataColumn | cell iteration benchmark | C# performance test | worksheet iteration speed | CheckCell method | .NET Excel processing
// Common Searches: Aspose.Cells benchmark MaxDataRow | measure cell iteration time C# | performance impact of MaxDataColumn | iterate over populated cells only Aspose | fixed range vs dynamic limits Excel library
// Developer Intent: Compare the execution time of cell loops that respect MaxDataRow/MaxDataColumn limits with loops that scan a larger, fixed range.
// Use Cases: Determine whether using MaxDataRow/MaxDataColumn reduces CPU time for large worksheets. | Validate that dynamic limits correctly bound iteration to the actual data region. | Create performance baselines for different worksheet sizes before optimizing data‑processing logic.
// AI Prompts: Write a reusable C# method that returns the elapsed milliseconds for iterating over cells with MaxDataRow and MaxDataColumn in Aspose.Cells. | Suggest an alternative loop that skips empty cells without calling CheckCell and report its speed gain. | Generate a summary table that logs iteration times for multiple row/column counts using both dynamic limits and a fixed oversized range.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMaxDataLimitsDemo
{
    // This example creates a workbook with 5,000 rows and 100 columns, saves it, and measures the time required to read all populated cells using the dynamic MaxDataRow/MaxDataColumn limits. It then repeats the read operation over a deliberately larger range (10,000 × 200) using CheckCell, allowing a direct performance comparison between constrained and unconstrained iteration.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the size of the test data
            int totalRows = 5000;      // number of rows to populate
            int totalColumns = 100;    // number of columns to populate

            // Populate the worksheet with sample data
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalColumns; col++)
                {
                    cells[row, col].PutValue($"R{row}_C{col}");
                }
            }

            // Save the workbook (uses the required save rule)
            workbook.Save("MaxDataLimitsDemo.xlsx");

            // -----------------------------------------------------------------
            // Measure iteration time using MaxDataRow / MaxDataColumn as limits
            // -----------------------------------------------------------------
            Stopwatch sw = new Stopwatch();
            sw.Start();

            int maxRow = cells.MaxDataRow;          // dynamic limit based on actual data
            int maxCol = cells.MaxDataColumn;       // dynamic limit based on actual data

            // Iterate only over the range that actually contains data
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    // Access the cell value (no operation needed, just read)
                    var value = cells[row, col].StringValue;
                }
            }

            sw.Stop();
            Console.WriteLine($"Iteration with MaxDataRow/MaxDataColumn limits: {sw.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Measure iteration time using a fixed larger range (without constraints)
            // -----------------------------------------------------------------
            Stopwatch sw2 = new Stopwatch();
            sw2.Start();

            // Define a larger range that exceeds the actual data size
            int fixedRows = 10000;   // intentionally larger than totalRows
            int fixedCols = 200;     // intentionally larger than totalColumns

            for (int row = 0; row < fixedRows; row++)
            {
                for (int col = 0; col < fixedCols; col++)
                {
                    // Access the cell; cells outside the populated area will be null or empty
                    var cell = cells.CheckCell(row, col);
                    if (cell != null)
                    {
                        var value = cell.StringValue;
                    }
                }
            }

            sw2.Stop();
            Console.WriteLine($"Iteration with fixed larger range: {sw2.ElapsedMilliseconds} ms");
        }
    }
}
