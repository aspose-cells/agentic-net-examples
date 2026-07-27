// Title: C# – Convert HTML to PDF with Aspose.Cells and embed Title & Author metadata
// Description: Loads an HTML file into an Aspose.Cells Workbook, sets the built‑in Title and Author properties, configures PdfSaveOptions to display the document title in PDF viewers, and saves the workbook as a PDF.
// Keywords: Aspose.Cells | C# | HTML to PDF conversion | PDF metadata | document title | author property | PdfSaveOptions | DisplayDocTitle | embed metadata | convert HTML workbook
// Common Searches: Aspose.Cells set PDF title metadata C# | How to add author to PDF when converting HTML with Aspose.Cells | DisplayDocTitle option example | Convert HTML file to PDF using Aspose.Cells C# | Set built‑in document properties before PDF export
// Developer Intent: The developer wants to convert an HTML file to PDF and embed title and author metadata in the resulting PDF.
// Use Cases: Generate PDF reports from HTML templates with proper metadata for document management systems. | Create downloadable PDFs from web pages where the viewer window shows the document title. | Automate batch conversion of HTML files to PDFs while preserving author information for compliance.
// AI Prompts: Show how to add additional PDF metadata such as subject, keywords, and creator using Aspose.Cells in C#. | Provide a loop example that converts multiple HTML files to PDFs, assigning a unique title and author to each file.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for PdfSaveOptions

// Loads an HTML file into an Aspose.Cells Workbook, sets the built‑in Title and Author properties, configures PdfSaveOptions to display the document title in PDF viewers, and saves the workbook as a PDF.
class HtmlToPdfWithMetadata
{
    static void Main()
    {
        // Load the HTML file into a workbook
        Workbook workbook = new Workbook("input.html");

        // Set built‑in document properties (title and author)
        workbook.BuiltInDocumentProperties.Title = "Sample Document Title";
        workbook.BuiltInDocumentProperties.Author = "John Doe";

        // Create PDF save options and enable title display in the PDF viewer
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.DisplayDocTitle = true; // ensures the PDF window title shows the document title

        // Save the workbook as a PDF file with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
