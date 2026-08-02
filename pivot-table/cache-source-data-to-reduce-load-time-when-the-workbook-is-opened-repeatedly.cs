using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // Optional: define a folder for temporary cache files used by Aspose.Cells
        // -----------------------------------------------------------------
        string cacheFolder = Path.Combine(Path.GetTempPath(), "AsposeCellsCache");
        Directory.CreateDirectory(cacheFolder);
        CellsHelper.SetCacheFolder(cacheFolder);

        // -----------------------------------------------------------------
        // Create a new workbook and fill it with a relatively large data set
        // -----------------------------------------------------------------
        Workbook workbook = new Workbook();                     // create workbook
        Worksheet sheet = workbook.Worksheets[0];              // first worksheet
        Cells cells = sheet.Cells;

        // Populate 10,000 rows with sample values
        for (int row = 0; row < 10000; row++)
        {
            cells[row, 0].PutValue(row);                      // numeric value
            cells[row, 1].PutValue($"Row {row}");             // text value
        }

        // -----------------------------------------------------------------
        // Start an access cache session.
        // We cache cell values (CellsData) and display‑related information (CellDisplay)
        // because the following loop reads both repeatedly.
        // -----------------------------------------------------------------
        workbook.StartAccessCache(AccessCacheOptions.CellsData | AccessCacheOptions.CellDisplay);

        // Access the data many times – the cache makes these operations faster
        for (int row = 0; row < 10000; row++)
        {
            int value = cells[row, 0].IntValue;               // fast read of cell value
            var style = cells[row, 0].GetDisplayStyle();      // fast read of display style
            // (no modifications are performed while the cache is active)
        }

        // -----------------------------------------------------------------
        // Close the cache to release resources and return to normal mode
        // -----------------------------------------------------------------
        workbook.CloseAccessCache(AccessCacheOptions.CellsData | AccessCacheOptions.CellDisplay);

        // -----------------------------------------------------------------
        // Save the workbook (uses the standard Save method – no custom code)
        // -----------------------------------------------------------------
        workbook.Save("CachedWorkbook.xlsx");
    }
}