// Title: Convert HTML String to PDF with Embedded Fonts using Aspose.Cells for .NET (C#)
// Description: Loads an HTML string from a MemoryStream into an Aspose.Cells Workbook and saves it as a PDF with embedded TrueType fonts (Arial) using PdfSaveOptions (EmbedStandardWindowsFonts, Identity encoding).
// Keywords: Aspose.Cells | C# | HTML to PDF | MemoryStream | HtmlLoadOptions | PdfSaveOptions | embed fonts | Identity encoding | Arial | .NET | convert HTML string | PDF generation | embedded fonts | workbook from stream
// Common Searches: Aspose.Cells convert HTML string to PDF C# | embed fonts in PDF using Aspose.Cells | load HTML from MemoryStream Aspose.Cells | PdfSaveOptions embed standard Windows fonts | HTML to PDF without temporary files Aspose.Cells
// Developer Intent: Create a PDF from an in‑memory HTML snippet while ensuring the output embeds the required TrueType fonts.
// Use Cases: Generate PDF reports from dynamic HTML content without writing intermediate files. | Maintain exact typography in PDFs for branding by embedding Arial or other Windows fonts. | Process multiple HTML fragments in batch, converting each to a font‑embedded PDF. | Integrate HTML‑to‑PDF conversion into web services or APIs.
// AI Prompts: Provide C# code that reads an HTML string into a MemoryStream, loads it into an Aspose.Cells Workbook with HtmlLoadOptions, and saves it as a PDF with embedded TrueType fonts using PdfSaveOptions. | Explain step‑by‑step how to configure PdfSaveOptions to embed standard Windows fonts, set Identity encoding, and define a fallback font when converting HTML to PDF with Aspose.Cells. | Show how to convert a list of HTML strings to PDFs in a loop, ensuring each PDF embeds the same fonts.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an HTML string from a MemoryStream into an Aspose.Cells Workbook and saves it as a PDF with embedded TrueType fonts (Arial) using PdfSaveOptions (EmbedStandardWindowsFonts, Identity encoding).
class HtmlToPdfWithEmbeddedFonts
{
    static void Main()
    {
        // HTML content to be loaded
        string html = "<html><body><p style='font-family:Arial; font-size:14pt;'>Hello, Aspose.Cells!</p></body></html>";

        // Convert the HTML string to a UTF‑8 byte array
        byte[] htmlBytes = System.Text.Encoding.UTF8.GetBytes(html);

        // Load the HTML from a memory stream using HtmlLoadOptions
        using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
        {
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(); // creates loading options
            // Example of optional settings:
            // loadOptions.SupportDivTag = true;
            // loadOptions.AutoFitColsAndRows = true;

            // Load the HTML into a Workbook
            Workbook workbook = new Workbook(htmlStream, loadOptions);

            // Configure PDF save options to embed fonts
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.EmbedStandardWindowsFonts = true;          // embed TrueType fonts
            pdfOptions.FontEncoding = PdfFontEncoding.Identity;   // use Identity encoding
            pdfOptions.DefaultFont = "Arial";                     // fallback font
            pdfOptions.CheckWorkbookDefaultFont = true;           // use workbook default font first

            // Save the workbook as PDF to a memory stream, then write to a file
            using (MemoryStream pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, pdfOptions);
                File.WriteAllBytes("output.pdf", pdfStream.ToArray());
            }
        }
    }
}
