using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data (optional, just to have printable content)
        worksheet.Cells["A1"].PutValue("Sample Header");
        worksheet.Cells["A2"].PutValue("Row 1");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Row 2");
        worksheet.Cells["B3"].PutValue(200);

        // Set page orientation to Landscape
        worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Fit the worksheet to 1 page wide; height will adjust automatically
        worksheet.PageSetup.FitToPagesWide = 1;
        worksheet.PageSetup.FitToPagesTall = 0; // 0 means auto‑scale height

        // Configure PDF save options to keep the layout as a single page per sheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true,
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook as a printable PDF
        workbook.Save("PrintableLandscape.pdf", pdfOptions);
    }
}