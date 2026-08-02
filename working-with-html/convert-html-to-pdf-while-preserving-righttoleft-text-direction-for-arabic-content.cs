// Title: Convert Arabic HTML to PDF with RTL Layout Using Aspose.Cells for .NET
// Description: Loads an Arabic HTML file into an Aspose.Cells workbook, activates right‑to‑left display, applies a TextDirection style to the used range, sets an Arabic‑compatible default font, and saves the result as a PDF that retains the RTL layout.
// Keywords: Aspose.Cells | HTML to PDF conversion | RTL layout | Arabic PDF | C# | DisplayRightToLeft | TextDirection | PdfSaveOptions | default font Arabic | spreadsheet to PDF
// Common Searches: Aspose.Cells convert Arabic HTML to PDF | C# preserve right‑to‑left text when exporting PDF | set default Arabic font in Aspose.Cells PDF output | apply RTL style to worksheet before PDF save | troubleshoot missing Arabic characters in PDF
// Developer Intent: Generate a PDF from an Arabic HTML source while keeping the right‑to‑left text direction intact.
// Use Cases: Produce printable PDFs of Arabic web reports with correct RTL alignment. | Create multilingual invoices where Arabic sections require RTL rendering before PDF export. | Automate batch conversion of Arabic HTML templates to PDFs that preserve layout and font rendering.
// AI Prompts: Write C# code that loads an Arabic HTML file into an Aspose.Cells workbook, applies a right‑to‑left text direction style to all cells, and saves it as a PDF with an Arabic‑compatible default font. | Explain how DisplayRightToLeft and TextDirection properties influence PDF rendering in Aspose.Cells and how to resolve missing Arabic glyphs. | Provide a step‑by‑step guide to batch‑process multiple HTML files into PDFs that retain RTL formatting using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using CellsRange = Aspose.Cells.Range;

namespace HtmlToPdfRtlExample
{
    // Loads an Arabic HTML file into an Aspose.Cells workbook, activates right‑to‑left display, applies a TextDirection style to the used range, sets an Arabic‑compatible default font, and saves the result as a PDF that retains the RTL layout.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source HTML file containing Arabic text
            string htmlPath = "input.html";

            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file \"{htmlPath}\" was not found.");
                return;
            }

            try
            {
                // Load the HTML file into a workbook.
                // Aspose.Cells can interpret HTML as a spreadsheet.
                Workbook workbook = new Workbook(htmlPath);

                // Iterate through all worksheets and enable right‑to‑left display.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Makes the whole sheet render from right to left.
                    sheet.DisplayRightToLeft = true;

                    // Create a style that forces right‑to‑left text direction.
                    Style rtlStyle = workbook.CreateStyle();
                    rtlStyle.TextDirection = TextDirectionType.RightToLeft;

                    // Apply the style to the used range of the worksheet.
                    // StyleFlag specifies which style attributes to apply.
                    StyleFlag flag = new StyleFlag
                    {
                        TextDirection = true
                    };

                    // MaxDisplayRange returns the area that actually contains data.
                    CellsRange usedRange = sheet.Cells.MaxDisplayRange;
                    usedRange.ApplyStyle(rtlStyle, flag);
                }

                // Configure PDF save options.
                // Set a default font that supports Arabic characters (e.g., Arial or Times New Roman).
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    DefaultFont = "Arial",
                    CheckWorkbookDefaultFont = true
                };

                // Save the workbook as a PDF file while preserving RTL layout.
                string pdfPath = "output.pdf";
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"HTML has been converted to PDF with RTL support: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
