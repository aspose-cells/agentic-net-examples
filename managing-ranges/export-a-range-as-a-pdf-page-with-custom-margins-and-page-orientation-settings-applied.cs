using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace ExportRangeToPdf
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (A1:C10)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the range to export
            string exportRange = "A1:C10";

            // Set the print area to the defined range
            sheet.PageSetup.PrintArea = exportRange;

            // Apply custom margins (in inches)
            sheet.PageSetup.LeftMarginInch = 0.5f;   // 0.5 inch left margin
            sheet.PageSetup.RightMarginInch = 0.5f;  // 0.5 inch right margin
            sheet.PageSetup.TopMarginInch = 0.75f;   // 0.75 inch top margin
            sheet.PageSetup.BottomMarginInch = 0.75f; // 0.75 inch bottom margin

            // Set page orientation (Landscape)
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Configure PDF save options to fit the range on a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true,               // Ensure one page per sheet
                AllColumnsInOnePagePerSheet = true    // Fit all columns into the page width
            };

            // Save the workbook as PDF; only the defined print area will be exported
            workbook.Save("ExportedRange.pdf", pdfOptions);

            Console.WriteLine("Range exported to PDF with custom margins and orientation.");
        }
    }
}