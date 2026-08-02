using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some content
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF version set to 1.6 (Pdf16)");

        // Configure PDF save options with compliance level Pdf16 (PDF 1.6)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Author note: setting compliance to Pdf16 enables PDF 1.6 features such as transparency
            Compliance = PdfCompliance.Pdf16
        };

        // Save the workbook as PDF using the configured options
        workbook.Save("OutputPdf16.pdf", pdfOptions);
    }
}