using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for ImageOrPrintOptions if needed

// Author: Aspose.Cells .NET example – set OnePagePerSheet and fit all columns on one PDF page
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the sheet with sample data that spans many columns
        for (int col = 0; col < 30; col++)
        {
            sheet.Cells[0, col].PutValue($"Header {col + 1}");
            for (int row = 1; row <= 20; row++)
            {
                sheet.Cells[row, col].PutValue($"R{row}C{col + 1}");
            }
        }

        // Configure page setup to fit all columns on a single page (rows may span multiple pages)
        PageSetup pageSetup = sheet.PageSetup;
        pageSetup.FitToPagesTall = 0;   // 0 = unlimited rows, forces all columns onto one page width
        pageSetup.FitToPagesWide = 1;   // optional: ensure only one page width is used

        // Create PDF save options and enable OnePagePerSheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true,
            // AllColumnsInOnePagePerSheet also achieves the same effect; included for completeness
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook as PDF with the configured options
        workbook.Save("Output.pdf", pdfOptions);

        Console.WriteLine("PDF generated with OnePagePerSheet = true and all columns on one page.");
    }
}