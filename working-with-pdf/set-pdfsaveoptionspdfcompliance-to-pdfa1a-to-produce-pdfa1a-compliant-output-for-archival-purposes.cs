using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfA1aExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in‑memory Excel file)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("PDF/A‑1a compliance example");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set the compliance level to PDF/A‑1a for archival purposes
            pdfOptions.Compliance = PdfCompliance.PdfA1a;

            // Save the workbook as a PDF file with the specified compliance level
            workbook.Save("PdfA1aOutput.pdf", pdfOptions);

            Console.WriteLine("PDF saved with PDF/A‑1a compliance.");
        }
    }
}