using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class SetPdfVersionExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF version 1.6 example");

        // Configure PDF save options to use PDF 1.6 compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.Pdf16 // Set PDF version to 1.6
        };

        // Save the workbook as a PDF file with the specified options
        workbook.Save("OutputPdf16.pdf", pdfOptions);
    }
}