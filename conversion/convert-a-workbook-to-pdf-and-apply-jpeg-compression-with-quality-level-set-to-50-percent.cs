using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfConversion
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a new workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF conversion");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set JPEG quality to 50% for images embedded in the PDF.
            // Desired PPI can be set to a typical screen resolution (e.g., 96).
            // The second parameter defines JPEG quality (0‑100).
            pdfOptions.SetImageResample(96, 50);

            // Save the workbook as PDF using the configured options
            workbook.Save("ConvertedWorkbook.pdf", pdfOptions);
        }
    }
}