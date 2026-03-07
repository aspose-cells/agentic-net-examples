using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPrintAllColumnsOnePage
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill the worksheet with sample data (e.g., 30 columns and 20 rows)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 30; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // -------------------------------------------------
            // Option 1: Save as PDF with all columns on one page
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.OnePagePerSheet = true;               // all content of the sheet on one page
            pdfOptions.AllColumnsInOnePagePerSheet = true;   // force all columns onto that page
            workbook.Save("AllColumnsOnePage.pdf", pdfOptions);

            // -------------------------------------------------
            // Option 2: Send directly to a printer (optional)
            // -------------------------------------------------
            // Configure print options to fit all columns on a single page
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions();
            printOptions.OnePagePerSheet = true;
            printOptions.AllColumnsInOnePagePerSheet = true;

            // Create a SheetRender instance after setting the options
            SheetRender renderer = new SheetRender(sheet, printOptions);

            // Replace "YourPrinterName" with the actual printer name installed on the system
            // renderer.ToPrinter("YourPrinterName");

            // Clean up resources
            renderer.Dispose();

            Console.WriteLine("PDF generated with all columns on a single page.");
        }
    }
}