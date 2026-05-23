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
            Workbook workbook = new Workbook();

            // Example data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample content for PDF conversion.");

            // Set the built‑in document title property
            workbook.BuiltInDocumentProperties.Title = "My Document Title";

            // Configure PDF save options to display the document title in the PDF viewer title bar
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DisplayDocTitle = true
            };

            // Save the workbook as a PDF file
            workbook.Save("Output.pdf", pdfOptions);
        }
    }
}