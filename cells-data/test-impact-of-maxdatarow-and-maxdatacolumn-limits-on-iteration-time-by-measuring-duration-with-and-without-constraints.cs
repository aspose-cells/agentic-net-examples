// Title: Measure Aspose.Cells C# worksheet iteration speed with and without MaxDataRow/MaxDataColumn limits
// Common Searches: how to benchmark cell iteration time using Cells.MaxDataRow in Aspose.Cells C# | compare performance of iterating all used cells versus a limited range in Aspose.Cells | C# code to time worksheet traversal with artificial row and column limits Aspose.Cells | measure impact of MaxDataRow and MaxDataColumn constraints on iteration speed in .NET | Aspose.Cells performance test for reading a large Excel worksheet in C#
// Tags: Aspose.Cells cell iteration performance | benchmark Cells.MaxDataRow traversal | measure worksheet traversal time C# | artificial row column limits performance Aspose.Cells | reading large Excel sheet timing Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMaxDataLimitsTest
{
    // The example creates a workbook with 10,000 rows and 100 columns, fills it with sample data, then uses Stopwatch to record the time required to iterate over all used cells via Cells.MaxDataRow/MaxDataColumn. It repeats the traversal with artificial limits set to half the rows and columns, reports both durations, and saves the file as MaxDataLimitsTest.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define the size of the test data
                const int totalRows = 10000;
                const int totalColumns = 100;

                // Populate the worksheet with sample data
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = 0; col < totalColumns; col++)
                    {
                        cells[row, col].PutValue($"R{row}C{col}");
                    }
                }

                // -----------------------------------------------------------------
                // Measure iteration time without any artificial limits
                // -----------------------------------------------------------------
                Stopwatch swNoLimits = Stopwatch.StartNew();

                // Iterate through all used cells
                for (int row = 0; row <= cells.MaxDataRow; row++)
                {
                    for (int col = 0; col <= cells.MaxDataColumn; col++)
                    {
                        // Access the cell value (no operation needed, just read)
                        var value = cells[row, col].Value;
                    }
                }

                swNoLimits.Stop();
                Console.WriteLine($"Iteration without limits: {swNoLimits.ElapsedMilliseconds} ms");

                // ---------------------------------------------------------------
                // Apply artificial row/column limits and measure again
                // ---------------------------------------------------------------
                int limitedRows = totalRows / 2;      // 5000
                int limitedCols = totalColumns / 2;   // 50

                // Ensure limits do not exceed actual used range
                int maxRow = Math.Min(limitedRows, cells.MaxDataRow);
                int maxCol = Math.Min(limitedCols, cells.MaxDataColumn);

                Stopwatch swWithLimits = Stopwatch.StartNew();

                // Iterate respecting the artificial limits
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        var value = cells[row, col].Value;
                    }
                }

                swWithLimits.Stop();
                Console.WriteLine($"Iteration with limits (Rows={maxRow + 1}, Columns={maxCol + 1}): {swWithLimits.ElapsedMilliseconds} ms");

                // Optional: Save the workbook to verify data (not required for timing)
                workbook.Save("MaxDataLimitsTest.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
