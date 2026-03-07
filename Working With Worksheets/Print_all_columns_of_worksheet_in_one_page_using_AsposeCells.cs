using System;
using Aspose.Cells;

namespace AsposeCellsPrintAllColumnsOnePage
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill the worksheet with sample data (e.g., 50 columns and 20 rows)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 50; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Configure PDF save options to fit all columns on a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.OnePagePerSheet = true;               // All content of the sheet on one page
            pdfOptions.AllColumnsInOnePagePerSheet = true;   // Force all columns onto that page

            // Save the workbook as PDF; the resulting PDF will have all columns on one page
            workbook.Save("AllColumnsOnePage.pdf", pdfOptions);

            Console.WriteLine("PDF saved with all columns on a single page.");
        }
    }
}