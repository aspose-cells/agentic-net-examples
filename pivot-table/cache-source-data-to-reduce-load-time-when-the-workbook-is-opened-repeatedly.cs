// Title: How to cache workbook data with Aspose.Cells in C# to speed up repeated openings
// AI Prompts: Generate C# code that starts an AccessCache for a workbook, reads random cells, then closes the cache before saving. | Show how to use Aspose.Cells StartAccessCache and CloseAccessCache to improve load performance for large worksheets. | Explain the steps to enable full workbook caching with AccessCacheOptions.All and release resources afterwards.
// Common Searches: Aspose.Cells C# cache workbook for faster subsequent opens | StartAccessCache example for large Excel files in .NET | How does CloseAccessCache affect memory usage in Aspose.Cells | Using AccessCacheOptions.All to speed up read‑only operations in Aspose.Cells | Performance tips for repeated reading of big worksheets with Aspose.Cells
// Tags: full workbook caching Aspose.Cells | read‑only operations with cached workbook | cache all workbook data Aspose.Cells | reduce Excel load time in .NET | cache large worksheet data for repeated access

using System;
using Aspose.Cells;

namespace AsposeCellsCacheDemo
{
    // // Demonstrates creating a workbook, populating 10,000 rows, starting a full access cache, reading random cells, closing the cache, and saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided Workbook constructor)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a relatively large data set (e.g., 10,000 rows)
            for (int row = 0; row < 10000; row++)
            {
                cells[row, 0].PutValue(row);               // Simple numeric data
                cells[row, 1].PutValue($"Item {row}");     // Text data
            }

            // Start an access cache session for the whole workbook.
            // AccessCacheOptions.All tells Aspose.Cells to cache everything that can be cached.
            workbook.StartAccessCache(AccessCacheOptions.All);

            // Perform read‑only operations that benefit from the cache.
            // For demonstration we read a few random cells and output their values.
            Random rnd = new Random();
            for (int i = 0; i < 5; i++)
            {
                int r = rnd.Next(0, 10000);
                Cell cellNumber = cells[r, 0];
                Cell cellText   = cells[r, 1];

                Console.WriteLine($"Row {r}: Number={cellNumber.IntValue}, Text=\"{cellText.StringValue}\"");
            }

            // Close the cache to release resources and return to normal mode.
            workbook.CloseAccessCache(AccessCacheOptions.All);

            // Save the workbook using the provided Save method.
            workbook.Save("CachedWorkbook.xlsx");

            Console.WriteLine("Workbook saved as CachedWorkbook.xlsx");
        }
    }
}
