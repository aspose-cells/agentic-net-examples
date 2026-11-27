// Title: Generate a PDF from an HTML string with embedded fonts using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells convert HTML string to PDF with embedded fonts | How to embed Arial font when saving PDF from HTML using Aspose.Cells | Load HTML from MemoryStream into Aspose.Cells workbook for PDF export | Aspose.Cells PDFSaveOptions embed standard Windows fonts example | Create one‑page‑per‑sheet PDF from HTML using Aspose.Cells in .NET
// Tags: html-to-pdf conversion aspnet aspose.cells | embed-standard-windows-fonts pdfsaveoptions | loadoptions html format workbook creation | memory-stream workbook loading aspnet | one-page-per-sheet pdf export option

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace HtmlToPdfWithEmbeddedFonts
{
    // The sample creates a MemoryStream from an HTML string, loads it into an Aspose.Cells Workbook with LoadOptions set to HTML format, configures PdfSaveOptions to embed standard Windows fonts and to place each worksheet on a separate PDF page, and then saves the workbook as a PDF file.
    class Program
    {
        static void Main(string[] args)
        {
            // HTML content to be converted. Replace with any valid HTML string as needed.
            string htmlContent = @"
                <html>
                    <head>
                        <style>
                            body { font-family: 'Arial'; font-size: 14pt; }
                        </style>
                    </head>
                    <body>
                        <p>Hello, Aspose.Cells PDF conversion with embedded fonts!</p>
                    </body>
                </html>";

            // Convert the HTML string to a memory stream.
            byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlContent);
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                try
                {
                    // Load the HTML into a Workbook using LoadOptions for HTML format.
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                    Workbook workbook = new Workbook(htmlStream, loadOptions);

                    // Configure PDF save options to embed fonts.
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        // Embed all standard Windows fonts (including Arial).
                        EmbedStandardWindowsFonts = true,
                        // Ensure each worksheet is saved on a separate PDF page.
                        OnePagePerSheet = true
                    };

                    // Define output PDF file path.
                    string outputPdfPath = "HtmlConverted.pdf";

                    // Save the workbook as PDF with the specified options.
                    workbook.Save(outputPdfPath, pdfOptions);

                    Console.WriteLine($"PDF successfully generated at: {Path.GetFullPath(outputPdfPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("An error occurred during conversion:");
                    Console.WriteLine(ex.Message);
                }
            }
        }
    }
}
