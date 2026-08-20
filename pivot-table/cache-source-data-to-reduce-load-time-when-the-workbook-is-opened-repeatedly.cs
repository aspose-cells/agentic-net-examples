// Title: C# – Use Aspose.Cells Access Cache to Accelerate Re‑opening Large Excel Workbooks
// Description: This example shows how to configure a temporary cache folder with CellsHelper.SetCacheFolder, create a workbook with 10,000 rows, and then reload it using StartAccessCache (CellsData | CellDisplay). Repeated reads are served from the cache, dramatically cutting load time, after which the cache is closed with CloseAccessCache and the file saved.
// Keywords: Aspose.Cells C# cache folder | StartAccessCache | CloseAccessCache | CellsHelper.SetCacheFolder | large workbook performance | Excel read optimization .NET | AccessCacheOptions CellsData | AccessCacheOptions CellDisplay | temporary cache directory Aspose | reduce Excel load time
// Common Searches: Aspose.Cells set temporary cache directory .NET | How to speed up loading of large Excel files with Aspose.Cells | StartAccessCache example C# | Cache cell values and display styles Aspose.Cells | CloseAccessCache after processing workbook
// Developer Intent: Enable cell‑data and display‑style caching to minimise the time required when the same large workbook is opened multiple times.
// Use Cases: Define a custom folder for Aspose.Cells temporary files to avoid default system temp usage. | Activate an access cache before intensive read loops on a worksheet containing thousands of rows. | Read cell values and display properties repeatedly while the cache is active for faster execution. | Terminate the cache after processing and optionally persist any changes to a new file. | Clean up the temporary cache folder when the operation is complete.
// AI Prompts: Generate a C# snippet that creates a custom cache folder, starts an access cache for CellsData and CellDisplay, reads a column repeatedly, then closes the cache. | Explain the performance impact of using AccessCacheOptions.CellsData versus AccessCacheOptions.CellDisplay in Aspose.Cells. | Provide best‑practice steps for managing the temporary cache directory when processing large Excel workbooks with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCacheDemo
{
    // This example shows how to configure a temporary cache folder with CellsHelper.SetCacheFolder, create a workbook with 10,000 rows, and then reload it using StartAccessCache (CellsData | CellDisplay). Repeated reads are served from the cache, dramatically cutting load time, after which the cache is closed with CloseAccessCache and the file saved.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Set a folder for temporary cache files (optional but improves
            //    performance for large workbooks).
            // -----------------------------------------------------------------
            string tempCacheFolder = Path.Combine(Path.GetTempPath(), "AsposeCellsCache");
            Directory.CreateDirectory(tempCacheFolder);
            CellsHelper.SetCacheFolder(tempCacheFolder);
            Console.WriteLine($"Cache folder set to: {tempCacheFolder}");

            // -----------------------------------------------------------------
            // 2. Create a workbook and populate it with sample data.
            // -----------------------------------------------------------------
            Workbook wb = new Workbook();                     // create workbook
            Worksheet sheet = wb.Worksheets[0];               // access first worksheet
            Cells cells = sheet.Cells;

            // Fill 10,000 rows with random numbers
            Random rnd = new Random();
            for (int row = 0; row < 10000; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue(rnd.NextDouble() * 1000);
                }
            }

            // Save the workbook to disk (first time, no cache involved)
            string filePath = "LargeDataWorkbook.xlsx";
            wb.Save(filePath);
            Console.WriteLine($"Workbook saved to {filePath}");

            // -----------------------------------------------------------------
            // 3. Load the workbook again and use access cache to speed up data reads.
            // -----------------------------------------------------------------
            Workbook loadedWb = new Workbook(filePath);       // load workbook
            Worksheet loadedSheet = loadedWb.Worksheets[0];
            Cells loadedCells = loadedSheet.Cells;

            // Start cache for cell values and display-related data
            loadedWb.StartAccessCache(AccessCacheOptions.CellsData | AccessCacheOptions.CellDisplay);
            Console.WriteLine("Access cache started.");

            // Perform many read operations that benefit from caching
            double sum = 0;
            for (int i = 0; i < 10000; i++)
            {
                // Access the first column value repeatedly
                sum += loadedCells[i, 0].DoubleValue;

                // Also retrieve display style (e.g., font size) to show cache usage
                var style = loadedCells[i, 0].GetDisplayStyle();
                // (no modification, just reading)
            }
            Console.WriteLine($"Sum of first column values: {sum}");

            // Close the cache after operations are done
            loadedWb.CloseAccessCache(AccessCacheOptions.CellsData | AccessCacheOptions.CellDisplay);
            Console.WriteLine("Access cache closed.");

            // Save the workbook again (optional, demonstrates normal save flow)
            loadedWb.Save("LargeDataWorkbook_CachedAccess.xlsx");
            Console.WriteLine("Workbook saved after cached access.");

            // Clean up: remove temporary cache folder if desired
            // Directory.Delete(tempCacheFolder, true);
        }
    }
}
