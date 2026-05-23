using System;
using Aspose.Cells;

namespace AsposeCellsRangeToPdf
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (A1:D10)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the range to be exported
            string exportRange = "A1:D10";

            // Set the print area to the defined range
            sheet.PageSetup.PrintArea = exportRange;

            // Apply custom margins (in inches)
            sheet.PageSetup.LeftMarginInch = 0.5f;   // 0.5 inch left margin
            sheet.PageSetup.RightMarginInch = 0.5f;  // 0.5 inch right margin
            sheet.PageSetup.TopMarginInch = 0.75f;   // 0.75 inch top margin
            sheet.PageSetup.BottomMarginInch = 0.75f; // 0.75 inch bottom margin

            // Set page orientation (Landscape)
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Ensure the range fits on a single PDF page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as PDF; only the defined print area will be exported
            workbook.Save("RangeExport.pdf", pdfOptions);

            Console.WriteLine("Range exported to PDF with custom margins and orientation.");
        }
    }
}