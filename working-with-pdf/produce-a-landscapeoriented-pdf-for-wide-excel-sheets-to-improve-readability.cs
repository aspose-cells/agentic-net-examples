using System;
using Aspose.Cells;

class LandscapePdfGenerator
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with many columns to simulate a wide sheet
        const int totalColumns = 30;
        const int totalRows = 20;

        for (int col = 0; col < totalColumns; col++)
        {
            // Header row
            sheet.Cells[0, col].PutValue($"Header {col + 1}");

            // Data rows
            for (int row = 1; row <= totalRows; row++)
            {
                sheet.Cells[row, col].PutValue($"R{row}C{col + 1}");
            }
        }

        // Set page orientation to Landscape for better horizontal space
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Fit all columns onto a single page width; height adjusts automatically
        sheet.PageSetup.FitToPagesWide = 1;   // one page wide
        sheet.PageSetup.FitToPagesTall = 0;   // auto height

        // Configure PDF save options to keep the entire sheet on one page
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true,
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook as a landscape‑oriented PDF
        workbook.Save("WideSheet_Landscape.pdf", pdfOptions);
    }
}