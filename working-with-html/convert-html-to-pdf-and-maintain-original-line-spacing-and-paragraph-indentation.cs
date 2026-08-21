// Title: C# – Convert HTML to PDF with Aspose.Cells while preserving line spacing and paragraph indentation
// Description: This example demonstrates how to load an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions (DeleteRedundantSpaces = false), optionally inspect text‑box shapes for TextParagraphs, and save the workbook as a PDF with PdfSaveOptions (OnePagePerSheet = false). The workflow keeps the original whitespace, line spacing, and indentation intact.
// Keywords: Aspose.Cells | C# | .NET | HTML to PDF conversion | preserve line spacing | paragraph indentation | DeleteRedundantSpaces | HtmlLoadOptions | PdfSaveOptions | OnePagePerSheet | text box shape | TextParagraph | formatting retention
// Common Searches: Aspose.Cells keep original spacing when converting HTML to PDF | DeleteRedundantSpaces HtmlLoadOptions effect on whitespace | C# convert HTML with text boxes to PDF preserving indentation | PdfSaveOptions OnePagePerSheet false example | how to retain line breaks in Aspose.Cells HTML import
// Developer Intent: Generate a PDF from an HTML document using Aspose.Cells without losing line breaks, spacing, or paragraph indentation.
// Use Cases: Produce printable PDFs from HTML reports that contain formatted text boxes. | Batch‑process email templates or web pages into PDFs while maintaining exact layout. | Create compliance‑critical PDFs (e.g., invoices, legal forms) where spacing and indentation must match the source HTML.
// AI Prompts: Write C# code that converts an HTML file to PDF with Aspose.Cells and keeps all line spacing and indentation. | Explain why setting DeleteRedundantSpaces to false preserves whitespace during HTML import. | Show how to modify TextParagraph properties of shape.TextBody after loading HTML to fine‑tune spacing.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsHtmlToPdf
{
    // This example demonstrates how to load an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions (DeleteRedundantSpaces = false), optionally inspect text‑box shapes for TextParagraphs, and save the workbook as a PDF with PdfSaveOptions (OnePagePerSheet = false). The workflow keeps the original whitespace, line spacing, and indentation intact.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source HTML file
                string htmlPath = "input.html";

                // Verify that the HTML file exists
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                    return;
                }

                // Path for the resulting PDF file
                string pdfPath = "output.pdf";

                // Ensure the output directory exists
                string pdfDir = Path.GetDirectoryName(pdfPath) ?? string.Empty;
                if (!string.IsNullOrEmpty(pdfDir) && !Directory.Exists(pdfDir))
                {
                    Directory.CreateDirectory(pdfDir);
                }

                // Load the HTML content into a workbook with specific options
                HtmlLoadOptions loadOptions = new HtmlLoadOptions
                {
                    DeleteRedundantSpaces = false
                };

                Workbook workbook = new Workbook(htmlPath, loadOptions);

                // Preserve line spacing in text box shapes (if any)
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Process only shapes that contain a TextBody (e.g., text boxes)
                        if (shape.TextBody != null)
                        {
                            foreach (TextParagraph paragraph in shape.TextBody.TextParagraphs)
                            {
                                // Placeholder for custom line spacing logic if needed
                            }
                        }
                    }
                }

                // Save the workbook as PDF with desired options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = false
                };

                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"HTML has been converted to PDF and saved at: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
