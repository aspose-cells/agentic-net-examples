// Title: Convert HTML with External CSS to PDF using Aspise.Cells (.NET)
// Description: Loads an HTML file that references external CSS files into an Aspose.Cells Workbook, then exports it to PDF while preserving the original stylesheet formatting.
// Keywords: Aspose.Cells HTML to PDF | external CSS conversion | C# HTML PDF export | preserve stylesheet Aspose.Cells | .NET HTML to PDF example | load HTML with linked CSS
// Common Searches: Aspose.Cells convert HTML with linked CSS to PDF C# | preserve external stylesheet when exporting HTML to PDF | HTML to PDF conversion using Aspose.Cells .NET | load HTML file with <link> tags in Aspose.Cells
// Developer Intent: Generate a PDF from an HTML document that uses external CSS, ensuring the visual style is retained, via Aspose.Cells for .NET.
// Use Cases: Create printable PDFs from web pages that rely on separate style sheets. | Archive HTML email newsletters with linked CSS as styled PDFs. | Produce branded invoices or reports from HTML templates that include external CSS.
// AI Prompts: Write C# code to load an HTML file containing <link> tags to external CSS into an Aspose.Cells Workbook and save it as a styled PDF. | Explain how Aspose.Cells resolves and applies external CSS during HTML‑to‑PDF conversion, and list any known limitations. | Suggest robust error‑handling and batch‑processing patterns for converting multiple HTML files with linked stylesheets to PDF using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlToPdf
{
    // Loads an HTML file that references external CSS files into an Aspose.Cells Workbook, then exports it to PDF while preserving the original stylesheet formatting.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains <link> tags referencing external CSS files
            string htmlPath = @"C:\Input\sample.html";

            // Path where the resulting PDF will be saved
            string pdfPath = @"C:\Output\sample.pdf";

            try
            {
                // Verify that the input HTML file exists
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Input file not found: {htmlPath}");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(pdfPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the HTML file into a Workbook; Aspose.Cells parses the HTML and applies external CSS.
                Workbook workbook = new Workbook(htmlPath);

                // Save the workbook as PDF. Visual appearance, including styles from external CSS, is preserved.
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine("HTML has been successfully converted to PDF with original stylesheet rules.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
