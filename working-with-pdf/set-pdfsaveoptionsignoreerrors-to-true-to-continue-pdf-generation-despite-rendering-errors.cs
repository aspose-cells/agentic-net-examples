using System;
using Aspose.Cells;

namespace AsposeCellsPdfIgnoreErrorDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // Example: new Workbook("input.xlsx");

            // Instantiate PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set IgnoreError to true so rendering errors do not stop PDF generation
            pdfOptions.IgnoreError = true;

            // Save the workbook as a PDF using the configured options
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}