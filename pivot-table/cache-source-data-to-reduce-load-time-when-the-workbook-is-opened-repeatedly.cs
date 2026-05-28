using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Optional: set a folder for temporary cache files used by Aspose.Cells
        string cacheFolder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "AsposeCellsCache");
        System.IO.Directory.CreateDirectory(cacheFolder);
        CellsHelper.SetCacheFolder(cacheFolder);

        // Create a new workbook and obtain the first worksheet
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate a large range with data and simple styling.
        // This data will be accessed repeatedly, so caching can improve performance.
        for (int i = 0; i < 1000; i++)
        {
            cells[i, 0].PutValue(i);

            // Apply alternating font colors as an example of style usage.
            Style style = wb.CreateStyle();
            style.Font.Color = (i % 2 == 0) ? Color.Blue : Color.Green;
            cells[i, 0].SetStyle(style);
        }

        // Start the access cache for cell values and display‑related information.
        // CellsData caches raw cell values; CellDisplay caches style/display data.
        wb.StartAccessCache(AccessCacheOptions.CellsData | AccessCacheOptions.CellDisplay);

        // Perform many read operations that benefit from the cache.
        for (int i = 0; i < 1000; i++)
        {
            Cell cell = cells[i, 0];

            // Retrieve the numeric value (cached).
            double value = cell.DoubleValue;

            // Retrieve the display style (cached).
            Style displayStyle = cell.GetDisplayStyle();

            // Output occasional progress to the console.
            if (i % 200 == 0)
                Console.WriteLine($"Row {i}: Value={value}, FontColor={displayStyle.Font.Color}");
        }

        // Close the cache to release resources and return to normal access mode.
        wb.CloseAccessCache(AccessCacheOptions.CellsData | AccessCacheOptions.CellDisplay);

        // Save the workbook. Subsequent openings can reuse the same caching pattern
        // to reduce load time when the same data is accessed repeatedly.
        wb.Save("CachedWorkbook.xlsx");
    }
}