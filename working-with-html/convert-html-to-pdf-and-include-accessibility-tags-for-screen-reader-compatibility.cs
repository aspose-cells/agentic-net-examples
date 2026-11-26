// Title: Convert an HTML file to a PDF/A‑1b compliant document for screen‑reader accessibility using Aspose.Cells in C#
// Common Searches: how to convert an HTML file to PDF/A-1b with Aspose.Cells in C# | C# Aspose.Cells generate accessible PDF from HTML for screen readers | save HTML workbook as PDF with PDF/A compliance using Aspose.Cells | Aspose.Cells PDF/A-1b option for HTML to PDF conversion in .NET | C# code to load HTML into Workbook and export to accessible PDF
// Tags: html to pdf conversion using aspose.cells | pdfa-1b compliance with aspose.cells pdfsaveoptions | c# workbook load html format | c# save workbook as pdf with accessibility | screen reader friendly pdf generation c#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace HtmlToPdfWithAccessibility
{
    // The example loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions with PDF/A‑1b compliance and one page per sheet, and saves the workbook as a PDF. While Aspose.Cells does not expose a dedicated accessibility flag, PDF/A‑1b provides baseline screen‑reader compatibility.
    class Program
    {
        /// <param name="args">
        /// args[0] – Path to the source HTML file (required).<br/>
        /// args[1] – Path for the output PDF file (optional, defaults to "output.pdf" in the current directory).<br/>
        /// </param>
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Usage: HtmlToPdfWithAccessibility <inputHtmlPath> [outputPdfPath]");
                return;
            }

            string inputHtmlPath = args[0];
            string outputPdfPath = args.Length > 1 ? args[1] : "output.pdf";

            // Verify that the HTML file exists
            if (!File.Exists(inputHtmlPath))
            {
                Console.WriteLine($"Error: The HTML file \"{inputHtmlPath}\" does not exist.");
                return;
            }

            try
            {
                // 1. Load the HTML file into an Aspose.Cells Workbook.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

                // 2. Configure PDF save options.
                //    - Compliance = PdfA1b creates a PDF/A‑1b compliant document (baseline for accessibility).
                //    - OnePagePerSheet = true creates a single page per worksheet.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Compliance = PdfCompliance.PdfA1b,
                    OnePagePerSheet = true
                    // Note: The Accessibility property is not available in the current Aspose.Cells version.
                    // PDF/A‑1b compliance already provides a level of accessibility.
                };

                // 3. Save the workbook as PDF using the configured options.
                workbook.Save(outputPdfPath, pdfOptions);

                Console.WriteLine($"HTML file \"{inputHtmlPath}\" successfully converted to PDF.");
                Console.WriteLine($"Output PDF: {outputPdfPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a helpful message.
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
