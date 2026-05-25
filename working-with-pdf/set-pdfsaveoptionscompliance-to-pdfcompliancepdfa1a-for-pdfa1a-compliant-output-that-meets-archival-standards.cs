using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfComplianceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("PDF/A‑1a compliance demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Create PDF save options and set compliance to PDF/A‑1a
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a
            };

            // Save the workbook as a PDF with the specified compliance level
            workbook.Save("PdfA1aOutput.pdf", pdfOptions);

            Console.WriteLine("PDF saved with PDF/A‑1a compliance.");
        }
    }
}