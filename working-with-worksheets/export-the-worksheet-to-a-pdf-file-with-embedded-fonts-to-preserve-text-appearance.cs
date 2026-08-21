// Title: Export Worksheet to PDF with Embedded Fonts using Aspose.Cells (C#)
// Description: Demonstrates how to save an Aspose.Cells workbook as a PDF while embedding the fonts used in the worksheet. The example configures PdfSaveOptions (EmbedStandardWindowsFonts, DefaultFont, CheckWorkbookDefaultFont, FontEncoding) to preserve exact text appearance across devices and support Unicode characters.
// Keywords: Aspose.Cells | C# | PDF export | embed fonts | PdfSaveOptions | preserve text appearance | default font | Identity encoding | Unicode PDF | worksheet to PDF
// Common Searches: Aspose.Cells embed fonts PDF | C# export worksheet to PDF with embedded fonts | PdfSaveOptions EmbedStandardWindowsFonts example | preserve font styling in PDF using Aspose.Cells | set default font for PDF export Aspose.Cells .NET
// Developer Intent: Generate a PDF from a worksheet that embeds all used fonts to keep the visual layout unchanged.
// Use Cases: Creating client‑ready PDFs that look identical on any device | Archiving financial reports with corporate fonts embedded | Producing printable invoices where font substitution must be avoided | Generating multilingual PDFs that correctly render Unicode characters | Distributing PDFs to users who may not have the original fonts installed
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as a PDF with all fonts embedded and using Identity encoding. | Explain how EmbedStandardWindowsFonts, DefaultFont, CheckWorkbookDefaultFont, and FontEncoding affect the resulting PDF. | Show how to embed a custom TrueType font from a specific folder when exporting a worksheet to PDF with Aspose.Cells. | Provide a step‑by‑step tutorial for preserving text appearance in PDFs generated from Aspose.Cells workbooks.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace ExportWorksheetToPdfWithEmbeddedFonts
{
    // Demonstrates how to save an Aspose.Cells workbook as a PDF while embedding the fonts used in the worksheet. The example configures PdfSaveOptions (EmbedStandardWindowsFonts, DefaultFont, CheckWorkbookDefaultFont, FontEncoding) to preserve exact text appearance across devices and support Unicode characters.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to the worksheet
            worksheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            worksheet.Cells["A2"].PutValue("This PDF will have embedded fonts.");

            // Optionally set a custom font for the cells (ensure the font is installed or in a custom folder)
            Style style = worksheet.Cells["A1"].GetStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 12;
            worksheet.Cells["A1"].SetStyle(style);
            worksheet.Cells["A2"].SetStyle(style);

            // Configure PDF save options to embed fonts and preserve text appearance
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Embed standard Windows TrueType fonts (default is true)
                EmbedStandardWindowsFonts = true,

                // Set the default font to use when a cell's font is missing or unsupported
                DefaultFont = "Arial",

                // Ensure the workbook's default font is checked for Unicode characters
                CheckWorkbookDefaultFont = true,

                // Use Identity encoding for all embedded fonts
                FontEncoding = PdfFontEncoding.Identity
            };

            // Save the workbook as a PDF file with the specified options
            workbook.Save("WorksheetWithEmbeddedFonts.pdf", pdfOptions);

            Console.WriteLine("PDF saved successfully with embedded fonts.");
        }
    }
}
