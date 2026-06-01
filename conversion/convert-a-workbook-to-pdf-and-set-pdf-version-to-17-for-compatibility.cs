using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for PDF conversion");

        // Set PDF save options with PDF 1.7 compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.Pdf17
        };

        // Save the workbook as a PDF file using the specified options
        workbook.Save("output.pdf", pdfOptions);

        Console.WriteLine("Workbook converted to PDF with PDF 1.7 compliance.");
    }
}