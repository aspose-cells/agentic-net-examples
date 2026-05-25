using System;
using System.Diagnostics;
using Aspose.Cells;

namespace RichTextIterationBenchmark
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Parameters for the test data
            const int totalRows = 5000;      // number of rows
            const int totalCols = 10;        // number of columns
            const string baseText = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            // Populate the worksheet with rich‑text cells.
            // Each cell will contain the baseText where every 5th character is bold.
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    Cell cell = cells[row, col];
                    cell.PutValue(baseText);

                    // Apply bold formatting to every 5th character
                    for (int i = 0; i < baseText.Length; i += 5)
                    {
                        int length = Math.Min(5, baseText.Length - i);
                        cell.Characters(i, length).Font.IsBold = true;
                    }
                }
            }

            // -----------------------------------------------------------------
            // Benchmark 1: Iterate without using access cache
            // -----------------------------------------------------------------
            Stopwatch swNoCache = Stopwatch.StartNew();

            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    Cell cell = cells[row, col];
                    string text = cell.StringValue;

                    // Iterate through each character and read its formatting
                    for (int i = 0; i < text.Length; i++)
                    {
                        // Retrieve the font for the single character
                        Font font = cell.Characters(i, 1).Font;
                        // Example read – we just access IsBold to force the lookup
                        bool isBold = font.IsBold;
                        // (No modification is performed)
                    }
                }
            }

            swNoCache.Stop();
            Console.WriteLine($"Iteration without cache elapsed: {swNoCache.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Benchmark 2: Iterate using access cache (CellsData + CellDisplay)
            // -----------------------------------------------------------------
            // Start the cache – the worksheet data will be treated as read‑only
            workbook.StartAccessCache(AccessCacheOptions.CellsData | AccessCacheOptions.CellDisplay);

            Stopwatch swCache = Stopwatch.StartNew();

            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    Cell cell = cells[row, col];
                    string text = cell.StringValue;

                    for (int i = 0; i < text.Length; i++)
                    {
                        Font font = cell.Characters(i, 1).Font;
                        bool isBold = font.IsBold;
                    }
                }
            }

            swCache.Stop();

            // Close the cache to restore normal access mode
            workbook.CloseAccessCache(AccessCacheOptions.CellsData | AccessCacheOptions.CellDisplay);

            Console.WriteLine($"Iteration with cache elapsed: {swCache.ElapsedMilliseconds} ms");

            // Save the workbook (optional, demonstrates use of the provided save rule)
            workbook.Save("RichTextBenchmarkResult.xlsx");
        }
    }
}