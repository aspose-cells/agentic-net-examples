// Title: Benchmark the impact of CellDisplay AccessCache on iterating rich‑text character formatting in a large Aspose.Cells worksheet (C#)
// AI Prompts: Execute the sample C# program and capture the elapsed milliseconds for the loop that reads bold and italic font properties without any cache, then run the same loop after calling wb.StartAccessCache(AccessCacheOptions.CellDisplay) and output the performance difference. | Extend the benchmark to test different worksheet sizes (e.g., 1,000‑row, 5,000‑row, 10,000‑row) and record how the CellDisplay access cache scales the iteration time for cell.Characters().Font accesses.
// Common Searches: Aspose.Cells C# benchmark reading cell.Characters font properties with and without CellDisplay cache | How much faster does rich‑text iteration become when using StartAccessCache in a large worksheet | Performance test for Aspose.Cells rich text formatting access in a 5,000 row workbook | Measure the effect of AccessCacheOptions.CellDisplay on font retrieval speed in .NET spreadsheets
// Tags: benchmark cell characters access cache Aspose.Cells | rich text font retrieval performance C# | CellDisplay access cache usage Aspose.Cells | large worksheet iteration speed test | Aspose.Cells performance testing rich text

using System;
using System.Diagnostics;
using Aspose.Cells;

// The program creates a 5,000‑row by 10‑column workbook with rich‑text cells, then benchmarks two loops that read bold and italic font properties—once without caching and once using wb.StartAccessCache(AccessCacheOptions.CellDisplay)—and prints the elapsed milliseconds for each approach.
class RichTextIterationBenchmark
{
    static void Main()
    {
        // Parameters for a large worksheet
        const int totalRows = 5000;
        const int totalCols = 10;

        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate the worksheet with rich‑text cells
        for (int r = 0; r < totalRows; r++)
        {
            for (int c = 0; c < totalCols; c++)
            {
                // Example text with three parts
                string text = "Bold Italic Normal";
                cells[r, c].PutValue(text);

                // Apply rich‑text formatting to each part
                // "Bold" part
                cells[r, c].Characters(0, 4).Font.IsBold = true;
                // "Italic" part
                cells[r, c].Characters(5, 6).Font.IsItalic = true;
                // "Normal" part left unchanged
            }
        }

        // -----------------------------------------------------------------
        // Benchmark 1: Iterate through rich‑text portions without caching
        // -----------------------------------------------------------------
        Stopwatch sw = Stopwatch.StartNew();

        for (int r = 0; r < totalRows; r++)
        {
            for (int c = 0; c < totalCols; c++)
            {
                Cell cell = cells[r, c];
                // Retrieve the display style for each character range
                // Here we simply read the font properties to simulate work
                Font fontBold = cell.Characters(0, 4).Font;
                Font fontItalic = cell.Characters(5, 6).Font;
                // Access a property to ensure the objects are materialized
                bool isBold = fontBold.IsBold;
                bool isItalic = fontItalic.IsItalic;
            }
        }

        sw.Stop();
        Console.WriteLine($"Iteration without cache: {sw.ElapsedMilliseconds} ms");

        // -----------------------------------------------------------------
        // Benchmark 2: Iterate through rich‑text portions with AccessCache
        // -----------------------------------------------------------------
        // Start cache that optimizes display‑related queries (rich‑text formatting)
        wb.StartAccessCache(AccessCacheOptions.CellDisplay);

        Stopwatch swCache = Stopwatch.StartNew();

        for (int r = 0; r < totalRows; r++)
        {
            for (int c = 0; c < totalCols; c++)
            {
                Cell cell = cells[r, c];
                Font fontBold = cell.Characters(0, 4).Font;
                Font fontItalic = cell.Characters(5, 6).Font;
                bool isBold = fontBold.IsBold;
                bool isItalic = fontItalic.IsItalic;
            }
        }

        swCache.Stop();
        Console.WriteLine($"Iteration with CellDisplay cache: {swCache.ElapsedMilliseconds} ms");

        // Close the cache to restore normal access mode
        wb.CloseAccessCache(AccessCacheOptions.CellDisplay);

        // Save the workbook (optional, demonstrates lifecycle usage)
        wb.Save("RichTextBenchmark.xlsx");
    }
}
