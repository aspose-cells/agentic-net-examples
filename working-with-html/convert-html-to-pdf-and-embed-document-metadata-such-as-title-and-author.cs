// Title: Convert HTML to PDF with Aspose.Cells for .NET and embed PDF title/author
// Description: Loads an HTML file into an Aspose.Cells Workbook, sets the built‑in Title and Author properties, configures PdfSaveOptions (DisplayDocTitle and Standard custom‑properties export), and saves the workbook as a PDF with the metadata embedded.
// Keywords: Aspose.Cells | .NET | HTML to PDF conversion | PDF metadata | document title | author property | PdfSaveOptions | DisplayDocTitle | PdfCustomPropertiesExport | embed PDF properties
// Common Searches: Aspose.Cells convert HTML to PDF with metadata | set PDF title and author using Aspose.Cells .NET | DisplayDocTitle option Aspose.Cells example | export workbook properties to PDF Aspose.Cells | embed custom properties in PDF with Aspose.Cells | HTML to PDF Aspose.Cells tutorial USA | Aspose.Cells PDF metadata guide India
// Developer Intent: Generate a PDF from an HTML source and embed built‑in document properties such as title and author.
// Use Cases: Produce compliance‑ready PDF reports from HTML templates while preserving author information. | Create PDFs whose window title bar displays the document title for a clearer user experience. | Transfer both built‑in and custom workbook properties into the PDF Info dictionary for downstream processing.
// AI Prompts: Show how to add custom workbook properties before exporting to PDF with Aspose.Cells. | Generate code that reads title and author from a JSON file and applies them to the workbook prior to PDF conversion. | Explain how to configure PdfSaveOptions to embed XMP metadata in addition to the standard PDF Info dictionary using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering; // For PdfCustomPropertiesExport enum

// Loads an HTML file into an Aspose.Cells Workbook, sets the built‑in Title and Author properties, configures PdfSaveOptions (DisplayDocTitle and Standard custom‑properties export), and saves the workbook as a PDF with the metadata embedded.
class HtmlToPdfWithMetadata
{
    static void Main()
    {
        // Load the HTML file into a workbook
        // The constructor automatically detects the format based on the file extension
        Workbook workbook = new Workbook("input.html");

        // Set built‑in document properties that will be embedded into the PDF
        workbook.BuiltInDocumentProperties.Title = "Sample Document Title";
        workbook.BuiltInDocumentProperties.Author = "John Doe";

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Ensure the PDF window title bar displays the document title
        pdfOptions.DisplayDocTitle = true;

        // Export built‑in and custom properties to the PDF (standard Info dictionary)
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Save the workbook as a PDF file with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
