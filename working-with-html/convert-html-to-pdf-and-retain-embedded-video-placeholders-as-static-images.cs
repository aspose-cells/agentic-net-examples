// Title: C# – Convert HTML with Embedded Videos to PDF (static image placeholders) using Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells Workbook, treats embedded video players as WebExtension shapes, and saves the workbook as PDF with PdfSaveOptions (EmbedAttachments = false) so videos appear as static image thumbnails, keeping the PDF lightweight.
// Keywords: Aspose.Cells | HTML to PDF | C# | video placeholder | WebExtension shape | PdfSaveOptions | EmbedAttachments false | static image conversion | .NET | PDF export
// Common Searches: Aspose.Cells convert HTML to PDF C# | HTML video to PDF placeholder image Aspose | How to keep video thumbnails when converting HTML to PDF .NET | Disable OLE attachments in Aspose.Cells PDF export | Render WebExtension shape as image in PDF Aspose
// Developer Intent: Create a PDF from an HTML page that contains video elements, ensuring the videos are represented by static images rather than embedded media.
// Use Cases: Generate printable PDFs from web dashboards that include video tutorials, showing each video as a thumbnail. | Archive multimedia‑rich web pages as compact PDFs without embedding large video files. | Produce documentation from HTML tutorials where video tags are replaced by placeholder images for offline viewing.
// AI Prompts: Show how to change the size of the placeholder image generated for embedded videos during HTML‑to‑PDF conversion with Aspose.Cells. | Provide code that adds PDF compliance and image‑quality settings to PdfSaveOptions while preserving video placeholders as static images. | Explain how to replace a WebExtension shape with a custom image before saving the workbook to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHtmlToPdf
{
    // Loads an HTML file into an Aspose.Cells Workbook, treats embedded video players as WebExtension shapes, and saves the workbook as PDF with PdfSaveOptions (EmbedAttachments = false) so videos appear as static image thumbnails, keeping the PDF lightweight.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that may contain embedded video players
            string htmlFilePath = "input.html";

            // Load the HTML file into a Workbook.
            // Aspose.Cells parses the HTML and creates corresponding worksheet objects.
            Workbook workbook = new Workbook(htmlFilePath);

            // When a web video player is present in the HTML, Aspose.Cells represents it
            // as a WebExtension shape. During PDF conversion the shape is rendered as a
            // static image placeholder, which is the desired behavior.
            // Ensure that no OLE attachments are embedded in the PDF (default is false).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = false
            };

            // Save the workbook as a PDF file.
            string pdfOutputPath = "output.pdf";
            workbook.Save(pdfOutputPath, pdfOptions);

            Console.WriteLine($"HTML file '{htmlFilePath}' has been converted to PDF '{pdfOutputPath}'.");
        }
    }
}
