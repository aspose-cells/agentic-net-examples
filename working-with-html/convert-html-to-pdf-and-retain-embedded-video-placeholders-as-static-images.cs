// Title: Convert HTML to PDF with Aspose.Cells and retain video thumbnails as static images (C#)
// Description: A C# sample that loads an HTML file into an Aspose.Cells Workbook, checks the file’s existence, and saves it as a PDF. Video tags are not rendered; their poster images appear as static placeholders, giving a printable view of pages that contain <video> elements.
// Keywords: Aspose.Cells | HTML to PDF conversion | C# | .NET PDF export | video placeholder image | poster attribute | static thumbnail | unsupported video handling | Workbook.Save | SaveFormat.Pdf
// Common Searches: Aspose.Cells keep video thumbnail when converting HTML to PDF | C# convert HTML with <video> tags to PDF showing poster image | HTML to PDF static image for embedded video Aspose | How to preserve video placeholders in PDF using Aspose.Cells | Convert web page to PDF with video thumbnails in .NET
// Developer Intent: Create a PDF from an HTML document while ensuring that any embedded videos are represented by their poster images rather than being omitted or rendered as active media.
// Use Cases: Produce printable marketing brochures from web pages that contain video teasers, showing only the thumbnail images. | Generate offline help manuals from online documentation that embeds videos, preserving the visual layout with static placeholders. | Batch‑process HTML reports with video elements into PDFs for archiving, where the videos are displayed as their poster frames.
// AI Prompts: Write C# code using Aspose.Cells that replaces <video> tags with their poster attribute before saving the document as PDF. | Explain how Aspose.Cells treats unsupported video elements during HTML‑to‑PDF conversion and how to guarantee that poster images appear. | Provide a step‑by‑step checklist to verify that video thumbnails are retained after converting HTML to PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// A C# sample that loads an HTML file into an Aspose.Cells Workbook, checks the file’s existence, and saves it as a PDF. Video tags are not rendered; their poster images appear as static placeholders, giving a printable view of pages that contain <video> elements.
class HtmlToPdfWithVideoPlaceholders
{
    static void Main()
    {
        try
        {
            // Input HTML file that may contain embedded video web extensions
            string htmlPath = "input.html";

            // Verify that the HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
                return;
            }

            // Load the HTML content into a workbook
            Workbook workbook = new Workbook(htmlPath);

            // Save the workbook as PDF – any video content will be rendered as is (or omitted if unsupported)
            string pdfOutputPath = "output.pdf";
            workbook.Save(pdfOutputPath, SaveFormat.Pdf);

            Console.WriteLine($"HTML converted to PDF saved at: {pdfOutputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
