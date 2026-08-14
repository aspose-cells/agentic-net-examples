// Title: Aspose.Cells for .NET – Configure Unicode Font Settings to Render Supplementary Characters in PDF
// Description: Shows how to assign a global Unicode fallback font (e.g., MS Gothic) and set PdfSaveOptions (DefaultFont, CheckWorkbookDefaultFont, CheckFontCompatibility, FontEncoding=Identity) so emojis, mathematical symbols and other supplementary characters are rendered correctly when a workbook is saved as PDF.
// Keywords: Aspose.Cells | C# | PDF export | Unicode font | supplementary characters | emoji rendering | PdfSaveOptions | DefaultFont | FontEncoding Identity | font fallback | MS Gothic | workbook to PDF | Unicode support
// Common Searches: Aspose.Cells render emoji in PDF | set default Unicode font for PDF export Aspose.Cells .NET | supplementary Unicode characters PDF Aspose.Cells | font fallback for unsupported characters Aspose.Cells PDF | PdfSaveOptions Unicode support example
// Developer Intent: Configure font settings so that PDF output from Aspose.Cells correctly displays supplementary Unicode characters such as emojis and mathematical symbols.
// Use Cases: Export a spreadsheet containing emojis or special symbols to PDF with accurate visual representation. | Provide a global fallback font for any cell that lacks a compatible typeface during PDF conversion. | Enable Identity font encoding to cover the full Unicode range when saving workbooks as PDFs.
// AI Prompts: Generate C# code that sets a Unicode fallback font and enables Identity encoding in Aspose.Cells PdfSaveOptions for full Unicode PDF output. | Explain how to configure Aspose.Cells to render supplementary characters (emoji, math alphanumerics) when saving a workbook to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to assign a global Unicode fallback font (e.g., MS Gothic) and set PdfSaveOptions (DefaultFont, CheckWorkbookDefaultFont, CheckFontCompatibility, FontEncoding=Identity) so emojis, mathematical symbols and other supplementary characters are rendered correctly when a workbook is saved as PDF.
class ConfigureFontSettingsForPdf
{
    static void Main()
    {
        // Set a global Unicode font that supports supplementary characters (e.g., MS Gothic)
        FontConfigs.DefaultFontName = "MS Gothic";

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add text containing supplementary Unicode characters (emoji, mathematical alphanumerics, etc.)
        worksheet.Cells["A1"].PutValue("Unicode test: 😀 𝔘𝔫𝔦𝔠𝔬𝔡𝔢");

        // Configure PDF save options to use the default Unicode font
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Font used when a cell does not specify a compatible font
            DefaultFont = "MS Gothic",
            // Try workbook's default font first
            CheckWorkbookDefaultFont = true,
            // Ensure font compatibility checking is enabled for fallback substitution
            CheckFontCompatibility = true,
            // Use Identity encoding to support all Unicode characters
            FontEncoding = PdfFontEncoding.Identity
        };

        // Save the workbook as PDF with the configured font settings
        workbook.Save("UnicodeOutput.pdf", pdfOptions);
    }
}
