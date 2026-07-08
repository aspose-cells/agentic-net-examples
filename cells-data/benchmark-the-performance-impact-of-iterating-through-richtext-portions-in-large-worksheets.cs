using System;
using System.Diagnostics;
using Aspose.Cells;

namespace RichTextBenchmark
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define worksheet size
            const int totalRows = 2000;   // large number of rows
            const int totalCols = 10;     // number of columns

            // Populate cells with rich‑text: first 4 characters bold, rest normal
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    // Base text
                    string text = "BoldNormalText";

                    // Put the text into the cell
                    cells[row, col].PutValue(text);

                    // Apply bold formatting to the first 4 characters
                    cells[row, col].Characters(0, 4).Font.IsBold = true;
                }
            }

            // Optional: start access cache for cells data to reduce overhead of reading
            workbook.StartAccessCache(AccessCacheOptions.CellsData);

            // Benchmark: iterate through all cells and read the rich‑text formatting
            Stopwatch sw = Stopwatch.StartNew();

            int boldPortionCount = 0; // just to use the retrieved data

            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    // Get the cell
                    Cell cell = cells[row, col];

                    // Retrieve the characters collection length
                    int length = cell.Value?.ToString().Length ?? 0;

                    // Scan the characters to find bold portions
                    // (In this example we know the first 4 chars are bold)
                    for (int i = 0; i < length; i++)
                    {
                        // Access the font of each character
                        Font font = cell.Characters(i, 1).Font;

                        if (font.IsBold)
                        {
                            boldPortionCount++;
                        }
                    }
                }
            }

            sw.Stop();

            // Close the cache after reading
            workbook.CloseAccessCache(AccessCacheOptions.CellsData);

            // Output benchmark result
            Console.WriteLine($"Iterated {totalRows * totalCols} cells.");
            Console.WriteLine($"Bold character count found: {boldPortionCount}");
            Console.WriteLine($"Time elapsed: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (optional, just to follow lifecycle)
            workbook.Save("RichTextBenchmarkResult.xlsx");
        }
    }
}