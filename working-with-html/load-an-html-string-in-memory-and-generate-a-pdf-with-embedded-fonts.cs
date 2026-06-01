using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class HtmlToPdfWithEmbeddedFonts
{
    static void Main()
    {
        // Sample HTML content to be converted
        string htmlContent = @"
            <html>
                <body>
                    <table>
                        <tr>
                            <td style='font-family:Arial; font-size:14pt;'>
                                Hello, Aspose.Cells!
                            </td>
                        </tr>
                    </table>
                </body>
            </html>";

        // Convert the HTML string to a UTF‑8 memory stream
        using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
        {
            // Load the HTML into a Workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlStream, loadOptions);

            // Configure PDF save options to embed fonts
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure TrueType fonts are embedded (default is true)
                EmbedStandardWindowsFonts = true,

                // Use Identity encoding for full Unicode support
                FontEncoding = PdfFontEncoding.Identity,

                // Fallback font if a cell's font is missing
                DefaultFont = "Arial",

                // Try to use the workbook's default font first
                CheckWorkbookDefaultFont = true
            };

            // Save the workbook as PDF with embedded fonts
            using (MemoryStream pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, pdfOptions);

                // Write the PDF to a file (optional, for verification)
                File.WriteAllBytes("Result.pdf", pdfStream.ToArray());
                Console.WriteLine("PDF generated successfully with embedded fonts.");
            }
        }
    }
}