// Title: Convert an HTML file with base64‑encoded images to PDF/A‑1b using Aspose.Cells for .NET while preserving image resolution
// Common Searches: c# convert html file containing base64 images to pdf using Aspose.Cells | preserve original image quality when exporting html with embedded base64 images to pdf in .NET | Aspose.Cells HtmlLoadOptions handling of base64 images for pdf conversion | generate PDF/A‑1b from html that includes base64 encoded pictures in C# | maintain image resolution during html to pdf conversion with Aspose.Cells
// Tags: Aspose.Cells HtmlLoadOptions base64 image handling | PdfSaveOptions PDF/A-1b export C# | preserve image resolution HTML to PDF conversion | convert HTML to PDF using Aspose.Cells workbook | C# base64 embedded images PDF generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace HtmlToPdfWithBase64Images
{
    // The example loads an HTML document that contains base64‑encoded images into an Aspose.Cells Workbook via HtmlLoadOptions, sets PdfSaveOptions for PDF/A‑1b compliance, and saves the workbook as a PDF, ensuring the embedded images retain their original resolution.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputHtmlPath = "input.html";
            const string outputPdfPath = "output.pdf";

            try
            {
                // Verify that the input HTML file exists
                if (!File.Exists(inputHtmlPath))
                {
                    Console.Error.WriteLine($"Input file not found: {inputHtmlPath}");
                    return;
                }

                // Load the HTML file (base64 images are handled automatically)
                HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(inputHtmlPath, htmlLoadOptions);

                // Configure PDF save options (PDF/A‑1b compliance)
                PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
                {
                    Compliance = PdfCompliance.PdfA1b
                };

                // Save the workbook as PDF
                workbook.Save(outputPdfPath, pdfSaveOptions);

                Console.WriteLine($"HTML file \"{inputHtmlPath}\" successfully converted to PDF \"{outputPdfPath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during conversion:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
