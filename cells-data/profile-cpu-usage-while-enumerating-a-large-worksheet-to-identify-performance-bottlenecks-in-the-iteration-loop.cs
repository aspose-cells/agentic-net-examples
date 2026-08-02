using System;
using System.Collections;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceProfiling
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill a large range with sample data (e.g., 10000 rows x 50 columns)
            int totalRows = 10000;
            int totalCols = 50;
            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    cells[r, c].PutValue(r * totalCols + c);
                }
            }

            // Optional: enable cache for cell data access to see its impact
            // Comment out the following two lines to compare without caching
            workbook.StartAccessCache(AccessCacheOptions.CellsData);
            // Note: after enumeration, close the cache
            try
            {
                // Prepare a stopwatch to measure elapsed time
                Stopwatch sw = Stopwatch.StartNew();

                // Get the cells enumerator
                IEnumerator enumerator = cells.GetEnumerator();

                // Iterate through all instantiated cells
                while (enumerator.MoveNext())
                {
                    // Cast the current object to Cell
                    Cell cell = (Cell)enumerator.Current;

                    // Simulate some work – read the value (this is the part we profile)
                    var value = cell.Value;
                }

                sw.Stop();

                // Output the profiling result
                Console.WriteLine($"Enumerated {cells.CountLarge} cells in {sw.Elapsed.TotalSeconds:F2} seconds.");
            }
            finally
            {
                // Close the cache if it was started
                workbook.CloseAccessCache(AccessCacheOptions.CellsData);
            }

            // Save the workbook (demonstrates usage of the required save rule)
            workbook.Save("LargeWorksheetProfiled.xlsx");
        }
    }
}