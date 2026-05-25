using System;
using Aspose.Cells;

namespace AsposeCellsPdfTitleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Define a name for the workbook (could be any string you consider as the workbook's name)
            string workbookName = "SampleWorkbook";

            // Optionally set the first worksheet's name to the workbook name
            workbook.Worksheets[0].Name = workbookName;

            // Set the built‑in document Title property to the workbook name
            // (uses Workbook.BuiltInDocumentProperties.Title rule)
            workbook.BuiltInDocumentProperties.Title = workbookName;

            // Create PDF save options and enable displaying the document title in the PDF window title bar
            // (uses PdfSaveOptions.DisplayDocTitle rule)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DisplayDocTitle = true
            };

            // Save the workbook as a PDF with the specified options (lifecycle rule: save)
            workbook.Save("SampleWorkbook.pdf", pdfOptions);
        }
    }
}