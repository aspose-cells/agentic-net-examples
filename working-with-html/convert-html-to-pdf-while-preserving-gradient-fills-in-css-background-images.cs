// Title: C# – Convert HTML with CSS Gradient Backgrounds to PDF using AspNet Aspose.Cells
// Description: Load an HTML file that contains CSS gradient backgrounds into an Aspose.Cells Workbook, automatically translate the gradients into shape fills, and save the workbook as a PDF while preserving the visual gradient effects.
// Keywords: Aspose.Cells HTML to PDF | C# convert HTML gradient to PDF | preserve CSS gradients Aspose | SaveFormat.Pdf gradient support | Aspose.Cells workbook HTML import
// Common Searches: Aspose.Cells keep CSS gradients when exporting HTML to PDF | C# convert HTML with gradient background to PDF | HTML to PDF conversion preserving gradient fills Aspose | How to render CSS gradients in PDF using Aspose.Cells
// Developer Intent: Convert an HTML document that uses CSS gradient backgrounds into a PDF while retaining the gradient appearance, using Aspose.Cells for .NET.
// Use Cases: Generate PDF reports from web‑styled HTML templates that rely on gradient backgrounds. | Batch‑process multiple HTML files with gradient styling into archival PDFs. | Expose a REST endpoint that receives HTML content, renders it with Aspose.Cells, and returns a gradient‑preserving PDF.
// AI Prompts: Write C# code to loop through a folder of *.html files and convert each to a PDF with gradient fills using Aspose.Cells. | Explain how to detect unsupported CSS gradients in Aspose.Cells and replace them with solid colors before PDF export. | Show how to modify the gradient fill parameters of shapes after loading HTML but before saving the workbook as PDF.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlToPdf
{
    // Load an HTML file that contains CSS gradient backgrounds into an Aspose.Cells Workbook, automatically translate the gradients into shape fills, and save the workbook as a PDF while preserving the visual gradient effects.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains CSS gradient background images
            string htmlPath = "input.html";

            // Path for the resulting PDF file
            string pdfPath = "output.pdf";

            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: Input file not found at '{htmlPath}'.");
                return;
            }

            try
            {
                // Load the HTML file into a new workbook instance.
                // The constructor automatically parses the HTML and converts CSS background images (including gradients)
                // into corresponding shapes with gradient fills inside the workbook.
                Workbook workbook = new Workbook(htmlPath);

                // Save the workbook as PDF.
                // The PDF renderer preserves the gradient fills that were created from the CSS backgrounds.
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"HTML file '{htmlPath}' has been successfully converted to PDF with gradients preserved at '{pdfPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
