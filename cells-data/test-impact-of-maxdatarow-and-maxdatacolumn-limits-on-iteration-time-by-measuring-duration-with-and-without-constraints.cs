using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMaxDataIterationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define size of test data
            const int totalRows = 5000;
            const int totalCols = 100;

            // Populate the worksheet with sample data
            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    cells[r, c].PutValue(r * totalCols + c);
                }
            }

            // Ensure data is committed before measuring
            workbook.CalculateFormula();

            // -----------------------------------------------------------------
            // 1. Measure iteration using MaxDataRow/MaxDataColumn each loop
            // -----------------------------------------------------------------
            Stopwatch sw = new Stopwatch();
            sw.Start();

            for (int r = 0; r <= cells.MaxDataRow; r++)               // property accessed each iteration
            {
                for (int c = 0; c <= cells.MaxDataColumn; c++)       // property accessed each iteration
                {
                    // Access the cell value (no operation needed, just read)
                    var val = cells[r, c].Value;
                }
            }

            sw.Stop();
            long durationWithRepeatedProperties = sw.ElapsedMilliseconds;

            // -----------------------------------------------------------------
            // 2. Measure iteration using cached MaxDataRow/MaxDataColumn values
            // -----------------------------------------------------------------
            int cachedMaxRow = cells.MaxDataRow;      // property called once
            int cachedMaxCol = cells.MaxDataColumn;   // property called once

            sw.Restart();

            for (int r = 0; r <= cachedMaxRow; r++)
            {
                for (int c = 0; c <= cachedMaxCol; c++)
                {
                    var val = cells[r, c].Value;
                }
            }

            sw.Stop();
            long durationWithCachedValues = sw.ElapsedMilliseconds;

            // -----------------------------------------------------------------
            // 3. Measure iteration over a fixed range without using MaxData* at all
            // -----------------------------------------------------------------
            sw.Restart();

            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    var val = cells[r, c].Value;
                }
            }

            sw.Stop();
            long durationFixedRange = sw.ElapsedMilliseconds;

            // Output the measured durations
            Console.WriteLine($"Iteration with MaxDataRow/MaxDataColumn each access: {durationWithRepeatedProperties} ms");
            Console.WriteLine($"Iteration with cached MaxDataRow/MaxDataColumn values: {durationWithCachedValues} ms");
            Console.WriteLine($"Iteration over fixed range (no MaxData* usage): {durationFixedRange} ms");

            // Save the workbook (demonstrates usage of the required save rule)
            workbook.Save("MaxDataIterationDemo.xlsx");
        }
    }
}