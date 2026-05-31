using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate enough rows to generate multiple PDF pages (optional)
        for (int i = 0; i < 200; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Set the header text in the center section (section index 1)
        // This places the text at the top center of every printed page
        worksheet.PageSetup.SetHeader(1, "My Centered Header");

        // Align header margins with page margins (default is true, kept for clarity)
        worksheet.PageSetup.IsHFAlignMargins = true;

        // Create PDF save options (default options are sufficient for header rendering)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF; the header will appear centered on each page
        workbook.Save("CenteredHeader.pdf", pdfOptions);
    }
}