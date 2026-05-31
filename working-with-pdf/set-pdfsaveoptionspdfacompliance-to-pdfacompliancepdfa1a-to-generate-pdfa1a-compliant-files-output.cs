using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfA1aExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF/A‑1a Example");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Create PDF save options
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Set the compliance level to PDF/A‑1a
            saveOptions.Compliance = PdfCompliance.PdfA1a;

            // Save the workbook as a PDF with the specified compliance
            workbook.Save("PdfA1aOutput.pdf", saveOptions);

            Console.WriteLine("PDF saved with PDF/A‑1a compliance.");
        }
    }
}