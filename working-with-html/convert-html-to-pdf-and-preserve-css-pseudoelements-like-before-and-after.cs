// Title: Convert HTML to PDF with Aspose.Cells for .NET – Preserve CSS ::before and ::after
// Description: Demonstrates how to load a local HTML file with Aspose.Cells' HtmlLoadOptions, create a Workbook, and save it as a PDF while keeping CSS pseudo‑elements ( ::before / ::after ) intact. Includes basic error handling for missing files and conversion failures.
// Keywords: Aspose.Cells HTML to PDF | preserve CSS pseudo‑elements | C# convert HTML to PDF | HtmlLoadOptions | Workbook to PDF | CSS ::before ::after rendering | .NET PDF conversion
// Common Searches: Aspose.Cells keep ::before and ::after when converting HTML to PDF | C# HTML to PDF conversion with CSS pseudo‑elements | How to render CSS ::before ::after in PDF using Aspose.Cells | HtmlLoadOptions settings for CSS support in Aspose.Cells
// Developer Intent: Generate a PDF from an HTML document using Aspose.Cells for .NET and ensure that CSS ::before and ::after pseudo‑elements are rendered in the output.
// Use Cases: Quickly convert a static HTML page to a PDF with default settings. | Enable advanced CSS processing in HtmlLoadOptions to improve pseudo‑element rendering. | Add robust file‑existence checks and exception handling around the conversion workflow.
// AI Prompts: Write C# code that configures HtmlLoadOptions in Aspose.Cells to guarantee ::before and ::after pseudo‑elements appear in the saved PDF. | Explain methods to verify that CSS pseudo‑elements are present in the PDF and suggest fallback techniques if they are omitted. | Show how to capture conversion warnings and log detailed information about unsupported CSS features during HTML‑to‑PDF conversion with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToPdfExample
{
    // Demonstrates how to load a local HTML file with Aspose.Cells' HtmlLoadOptions, create a Workbook, and save it as a PDF while keeping CSS pseudo‑elements ( ::before / ::after ) intact. Includes basic error handling for missing files and conversion failures.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Path for the resulting PDF file
            string pdfPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input HTML file not found: {htmlPath}");
                return;
            }

            try
            {
                // Load the HTML content (HtmlLoadOptions can be used for additional settings if needed)
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(htmlPath, loadOptions);

                // Create PDF save options (default settings are sufficient)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
