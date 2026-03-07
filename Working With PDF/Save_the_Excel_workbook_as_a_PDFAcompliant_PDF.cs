using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfAExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF/A Example");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Configure PDF save options for PDF/A compliance (e.g., PDF/A-1b)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b
            };

            // Save the workbook as a PDF/A‑compliant PDF
            workbook.Save("Output_PdfA1b.pdf", pdfOptions);

            Console.WriteLine("Workbook saved as PDF/A‑1b compliant PDF.");
        }
    }
}