using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data (optional, just to see content in the PDF)
        worksheet.Cells["A1"].PutValue("Document with 0.5 inch margins on all sides");

        // Set all page margins to 0.5 inches using the Inch properties
        worksheet.PageSetup.LeftMarginInch   = 0.5;
        worksheet.PageSetup.RightMarginInch  = 0.5;
        worksheet.PageSetup.TopMarginInch    = 0.5;
        worksheet.PageSetup.BottomMarginInch = 0.5;

        // Create PDF save options (default options are sufficient for margin settings)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF file with the specified margins
        workbook.Save("DocumentWithCustomMargins.pdf", pdfOptions);
    }
}