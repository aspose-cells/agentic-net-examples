using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class SetPdfVersionExample
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF version 1.4 example");

        // Create PDF save options (rule: PdfSaveOptions constructor)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set the compliance level to PDF 1.4 (rule: PdfSaveOptions.Compliance property)
        pdfOptions.Compliance = PdfCompliance.Pdf14;

        // Save the workbook as a PDF using the specified options (rule: Workbook.Save)
        workbook.Save("Output_Pdf14.pdf", pdfOptions);
    }
}