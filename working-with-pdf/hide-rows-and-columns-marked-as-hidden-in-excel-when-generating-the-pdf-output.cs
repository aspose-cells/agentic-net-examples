using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class HideRowsColumnsPdfDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data
        worksheet.Cells["A1"].PutValue("Visible Row 1");
        worksheet.Cells["A2"].PutValue("Hidden Row");
        worksheet.Cells["A3"].PutValue("Visible Row 2");
        worksheet.Cells["B1"].PutValue("Visible Column 1");
        worksheet.Cells["B2"].PutValue("Hidden Column");
        worksheet.Cells["B3"].PutValue("Visible Column 2");

        // Hide the second row (index 1) and the second column (index 1)
        worksheet.Cells.Rows[1].IsHidden = true;      // Hide row 2
        worksheet.Cells.Columns[1].IsHidden = true;   // Hide column B

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Optional: ignore completely blank pages in the output PDF
            PrintingPageType = PrintingPageType.IgnoreBlank
        };

        // Save the workbook to PDF; hidden rows and columns will not appear in the PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}