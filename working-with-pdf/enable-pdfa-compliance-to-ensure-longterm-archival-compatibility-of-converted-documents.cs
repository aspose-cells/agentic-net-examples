using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfAComplianceExample
{
    static void Main()
    {
        // Create a new workbook and add some content
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF/A compliance example");

        // Set PDF save options to use PDF/A-1b compliance (ISO 19005-1)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the workbook as a PDF file with the specified compliance level
        workbook.Save("PdfA1b_Output.pdf", pdfOptions);
    }
}