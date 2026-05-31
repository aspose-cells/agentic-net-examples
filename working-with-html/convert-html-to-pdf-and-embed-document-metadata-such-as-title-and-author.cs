using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class HtmlToPdfWithMetadata
{
    static void Main()
    {
        // Load the HTML file into a workbook (create/load step)
        Workbook workbook = new Workbook("input.html");

        // Set built‑in document properties (metadata)
        workbook.BuiltInDocumentProperties.Title = "Sample Document Title";
        workbook.BuiltInDocumentProperties.Author = "Jane Smith";

        // Create PDF save options and enable title display in the PDF window title bar
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.DisplayDocTitle = true; // uses PdfSaveOptions.DisplayDocTitle property

        // Save the workbook as a PDF file with the specified options (save step)
        workbook.Save("output.pdf", pdfOptions);
    }
}