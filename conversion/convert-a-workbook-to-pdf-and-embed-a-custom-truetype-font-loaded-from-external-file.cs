// Title: Embed an External TrueType Font in PDF Export with Aspose.Cells (C#)
// Description: Registers a custom font folder, applies the TrueType font to worksheet cells, configures PdfSaveOptions (DefaultFont, CheckWorkbookDefaultFont, EmbedStandardWindowsFonts, FontEncoding) and saves the workbook as a PDF with the font fully embedded.
// Keywords: Aspose.Cells custom font PDF | C# embed TrueType font Aspose.Cells | PdfSaveOptions DefaultFont | FontConfigs.SetFontFolder | Unicode PDF generation Aspose.Cells | embed external font Aspose.Cells .NET | PDF export with custom typeface
// Common Searches: how to embed a custom TrueType font in Aspose.Cells PDF | set custom font folder for Aspose.Cells rendering | Aspose.Cells PdfSaveOptions use specific font | embed Unicode fonts in PDF with Aspose.Cells C# | Aspose.Cells export workbook to PDF with custom font
// Developer Intent: Include a TrueType font from an external directory in the PDF produced by Aspose.Cells, ensuring the font is embedded and displayed correctly.
// Use Cases: Create branded PDF reports where a corporate font must appear exactly as designed. | Deploy a shared font repository for multiple workbooks and guarantee consistent rendering without installing the font on each client machine. | Generate multilingual PDFs with full Unicode support by embedding the required font and using Identity encoding.
// AI Prompts: Write C# code that uses Aspose.Cells to export a workbook to PDF while embedding a TrueType font located in a custom folder. | Explain the steps to configure FontConfigs.SetFontFolder and PdfSaveOptions so that a custom font is embedded and Unicode characters render correctly. | Show how to set DefaultFont and CheckWorkbookDefaultFont in PdfSaveOptions to apply a custom font when cells have no explicit font definition.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCustomFontPdf
{
    // Registers a custom font folder, applies the TrueType font to worksheet cells, configures PdfSaveOptions (DefaultFont, CheckWorkbookDefaultFont, EmbedStandardWindowsFonts, FontEncoding) and saves the workbook as a PDF with the font fully embedded.
    class Program
    {
        static void Main()
        {
            // Path to the folder that contains the custom TrueType font file.
            // Set recursive to true to scan subfolders as well.
            FontConfigs.SetFontFolder(@"C:\CustomFonts", true);

            // Create a new workbook (or load an existing one).
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Put some sample text that uses the custom font.
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Text rendered with custom TrueType font");

            // Apply the custom font to the cell style.
            Style style = cell.GetStyle();
            style.Font.Name = "MyCustomFont"; // Replace with the actual font name inside the folder.
            style.Font.Size = 14;
            cell.SetStyle(style);

            // Configure PDF save options to embed the custom font.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Use the custom font as the default when the cell does not specify a font.
                DefaultFont = "MyCustomFont",

                // Try to use the workbook's default font first (helps with Unicode characters).
                CheckWorkbookDefaultFont = true,

                // Ensure the font is embedded in the PDF.
                EmbedStandardWindowsFonts = true,

                // Use Identity encoding for full Unicode support.
                FontEncoding = PdfFontEncoding.Identity
            };

            // Save the workbook as PDF with the specified options.
            workbook.Save(@"C:\Output\CustomFontDocument.pdf", pdfOptions);

            Console.WriteLine("PDF generated with embedded custom TrueType font.");
        }
    }
}
