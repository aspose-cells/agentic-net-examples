// Title: Create a PDF whose page size matches an Excel worksheet using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to export a workbook to PDF with physical dimensions that exactly equal the rendered size of the sheet. The example clears margins, forces a single‑page layout, retrieves width and height via SheetRender.GetPageSizeInch, applies CustomPaperSize, and saves the PDF.
// Keywords: Aspose.Cells | C# | custom paper size | SheetRender GetPageSizeInch | fit worksheet to one page | remove margins PDF export | Excel to PDF exact dimensions | Aspose.Cells PDF export | paper size custom | PDF page size from worksheet
// Common Searches: Aspose.Cells set custom PDF page size | How to match PDF page size to Excel sheet in .NET | Get worksheet rendered size in inches Aspose.Cells | Export Excel to PDF without scaling Aspose | C# Aspose.Cells custom paper size PDF
// Developer Intent: Export an Excel worksheet to PDF with a page size that mirrors the worksheet’s rendered dimensions.
// Use Cases: Printing reports that must retain the exact layout of the original sheet | Generating dashboards where each grid fits a single PDF page of a specific size | Creating printable forms that require no extra margins or scaling | Automating batch conversion of worksheets to size‑specific PDFs
// AI Prompts: Provide C# code using Aspose.Cells to calculate a worksheet’s rendered width and height and set those values as a custom PDF paper size. | Explain step‑by‑step how to clear margins, fit a sheet to one page, retrieve page dimensions with SheetRender, and export to PDF. | Show how to combine SheetRender.GetPageSizeInch with PageSetup.CustomPaperSize to produce a PDF that exactly matches the worksheet size.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCustomPdfSize
{
    // Demonstrates how to export a workbook to PDF with physical dimensions that exactly equal the rendered size of the sheet. The example clears margins, forces a single‑page layout, retrieves width and height via SheetRender.GetPageSizeInch, applies CustomPaperSize, and saves the PDF.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (adjust as needed)
                for (int row = 0; row < 20; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Define the print area and remove margins for exact sizing
                sheet.PageSetup.PrintArea = "A1:E20";
                sheet.PageSetup.LeftMargin = 0;
                sheet.PageSetup.RightMargin = 0;
                sheet.PageSetup.TopMargin = 0;
                sheet.PageSetup.BottomMargin = 0;

                // Fit the whole area onto a single page (helps SheetRender calculate size)
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 1;

                // Temporarily set paper size to Custom so that later we can assign exact dimensions
                sheet.PageSetup.PaperSize = PaperSizeType.Custom;

                // Use SheetRender to obtain the size (in inches) of the rendered page
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true
                    // ImageFormat is not required for size calculation
                };

                SheetRender sheetRender = new SheetRender(sheet, renderOptions);
                // Get width and height of the first (and only) page
                float[] pageSizeInInches = sheetRender.GetPageSizeInch(0);
                double pageWidth = pageSizeInInches[0];
                double pageHeight = pageSizeInInches[1];

                // Apply the exact dimensions as a custom paper size
                sheet.PageSetup.CustomPaperSize(pageWidth, pageHeight);

                // Save the workbook as PDF; the custom paper size will be used
                workbook.Save("CustomSizeOutput.pdf", SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
