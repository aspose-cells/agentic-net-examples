using System;
using Aspose.Cells;

namespace FitToPagesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet's PageSetup
            PageSetup pageSetup = workbook.Worksheets[0].PageSetup;

            // Benefit 1: Fit all columns to a single page width.
            // Setting FitToPagesWide = 1 ensures the printed sheet will not span multiple horizontal pages.
            pageSetup.FitToPagesWide = 1;

            // Benefit 2: Let the height adjust automatically.
            // Setting FitToPagesTall = 0 tells Excel to calculate the required number of pages vertically,
            // preventing content from being squeezed vertically.
            pageSetup.FitToPagesTall = 0;

            // Benefit 3: Consistent pagination across different printers.
            // By defining explicit page scaling, the output looks the same regardless of printer DPI.

            // Benefit 4: Simplifies printing of large tables.
            // Users can print the entire table on one sheet width, making it easier to read and file.

            // Save the workbook (save rule)
            workbook.Save("FitToPagesDemo.xlsx");
        }
    }
}