using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPrintFitDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data that spans multiple pages
            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Optional: define the print area
            sheet.PageSetup.PrintArea = "A1:J100";

            // Fit the worksheet to 1 page wide and 1 page tall when printed
            sheet.PageSetup.FitToPagesWide = 1;   // number of pages wide
            sheet.PageSetup.FitToPagesTall = 1;   // number of pages tall
            // Equivalent alternative:
            // sheet.PageSetup.SetFitToPages(1, 1);

            // Save the workbook (to verify the settings if needed)
            workbook.Save("FitPrintDemo.xlsx");

            // Prepare print options
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true   // each printed page corresponds to a separate sheet page
            };

            // Create a renderer for the worksheet
            SheetRender renderer = new SheetRender(sheet, printOptions);

            // Send the worksheet to a printer.
            // Replace "YourPrinterName" with an actual printer name, or pass null to use the default printer.
            // renderer.ToPrinter("YourPrinterName");
            // renderer.ToPrinter(null); // uses default printer

            // Clean up resources
            renderer.Dispose();
        }
    }
}