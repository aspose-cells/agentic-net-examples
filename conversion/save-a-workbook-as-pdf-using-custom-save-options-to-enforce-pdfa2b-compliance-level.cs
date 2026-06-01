using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfA2bSaveDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF/A-2b compliance demo");

        // Create PDF save options and set the compliance level to PDF/A-2b
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.Compliance = PdfCompliance.PdfA2b;

        // Save the workbook as a PDF file using the custom options
        workbook.Save("PdfA2bDemo.pdf", saveOptions);
    }
}