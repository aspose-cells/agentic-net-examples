// Title: C# – Convert Arabic HTML to PDF with RTL layout using Aspose.Cells
// Description: Loads an HTML file containing Arabic text into an Aspose.Cells workbook, enables right‑to‑left display, applies a RTL text direction style to the used range, sets an Arabic‑compatible font (e.g., Arial) via PdfSaveOptions, and saves the result as a PDF while preserving the correct RTL layout.
// Keywords: Aspose.Cells HTML to PDF | C# RTL conversion | Arabic PDF export | right to left layout Aspose.Cells | .NET convert HTML to PDF | DisplayRightToLeft property | TextDirection RightToLeft | PdfSaveOptions DefaultFont | Arabic font support
// Common Searches: Aspose.Cells convert Arabic HTML to PDF | C# export HTML with RTL to PDF using Aspose | How to keep right‑to‑left formatting in PDF with Aspose.Cells | Set default Arabic font for PDF export Aspose.Cells | Enable DisplayRightToLeft before saving PDF
// Developer Intent: Create a PDF from an Arabic HTML source while preserving right‑to‑left text direction and proper character rendering.
// Use Cases: Automated pipeline that converts web‑generated Arabic reports (HTML) to PDF for archiving. | Generating printable invoices or certificates in Arabic where layout must remain RTL. | Building a multilingual document conversion service that supports Arabic without manual post‑processing.
// AI Prompts: Generate C# code with Aspose.Cells to convert an Arabic HTML file to PDF, preserving RTL layout and using Arial as the default font. | Explain the impact of DisplayRightToLeft and TextDirection properties on RTL rendering when exporting to PDF with Aspose.Cells. | Provide a step‑by‑step troubleshooting guide for garbled Arabic characters after HTML‑to‑PDF conversion using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlToPdfRtl
{
    // Loads an HTML file containing Arabic text into an Aspose.Cells workbook, enables right‑to‑left display, applies a RTL text direction style to the used range, sets an Arabic‑compatible font (e.g., Arial) via PdfSaveOptions, and saves the result as a PDF while preserving the correct RTL layout.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source HTML file containing Arabic text
                string htmlPath = "input.html";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Input file not found: {htmlPath}");
                    return;
                }

                // Load the HTML file into a new workbook instance
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(htmlPath, loadOptions);

                // Access the first worksheet (or any specific worksheet)
                Worksheet worksheet = workbook.Worksheets[0];

                // Enable right‑to‑left display for the worksheet
                worksheet.DisplayRightToLeft = true;

                // Ensure that individual cells also have RTL direction
                Style rtlStyle = workbook.CreateStyle();
                rtlStyle.TextDirection = TextDirectionType.RightToLeft;

                // Apply the style to the used range
                Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;
                usedRange.ApplyStyle(rtlStyle, new StyleFlag { TextDirection = true });

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Use a font that supports Arabic characters
                    DefaultFont = "Arial",
                    // Try to use the workbook's default font first
                    CheckWorkbookDefaultFont = true
                };

                // Save the workbook as PDF while preserving RTL layout
                string outputPath = "output.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
