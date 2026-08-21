// Title: How to Embed All Fonts in a PDF Export with Aspose.Cells for .NET (C#)
// Description: Demonstrates creating a workbook, configuring PdfSaveOptions (EmbedStandardWindowsFonts, Identity encoding, default fallback font, workbook‑font checks) and saving it as a PDF that contains every font used, ensuring identical rendering on any device.
// Keywords: Aspose.Cells PDF font embedding | C# embed fonts Aspose.Cells | PdfSaveOptions EmbedStandardWindowsFonts | Identity font encoding Aspose.Cells | CheckFontCompatibility Aspose.Cells | custom font folder Aspose.Cells | export workbook to PDF with embedded fonts
// Common Searches: Aspose.Cells embed all fonts when saving to PDF | PdfSaveOptions font embedding C# example | How to ensure PDF uses embedded TrueType fonts in Aspose.Cells | Set default font and encoding for PDF export Aspose.Cells | Check font compatibility for multilingual PDF with Aspose.Cells
// Developer Intent: Include every font referenced in the workbook inside the generated PDF to guarantee consistent appearance.
// Use Cases: Produce PDF reports that preserve corporate or brand‑specific typefaces across all viewers. | Export workbooks that rely on custom fonts stored in a local directory, embedding them automatically. | Create multilingual PDFs where all glyphs are retained by validating font compatibility during conversion.
// AI Prompts: Generate C# code using Aspose.Cells that embeds all workbook fonts, including custom fonts from a folder, when saving to PDF. | Explain the impact of PdfSaveOptions properties such as EmbedStandardWindowsFonts, FontEncoding, and CheckFontCompatibility on PDF font embedding. | Provide a step‑by‑step tutorial for configuring Aspose.Cells to embed fonts and handle missing glyphs for international content.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates creating a workbook, configuring PdfSaveOptions (EmbedStandardWindowsFonts, Identity encoding, default fallback font, workbook‑font checks) and saving it as a PDF that contains every font used, ensuring identical rendering on any device.
class EmbedFontsToPdf
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample text with embedded fonts");

        // If you have custom fonts, point Aspose.Cells to the folder containing them
        // FontConfigs.SetFontFolder(@"C:\MyFonts", true);

        // Configure PDF save options to embed all fonts
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed TrueType fonts (default is true, set explicitly for clarity)
            EmbedStandardWindowsFonts = true,

            // Use Identity encoding to ensure all characters are embedded correctly
            FontEncoding = PdfFontEncoding.Identity,

            // Fallback font if a specific font is missing
            DefaultFont = "Arial",

            // Ensure workbook's default font is considered for Unicode characters
            CheckWorkbookDefaultFont = true,

            // Verify font compatibility for each character (helps avoid missing glyphs)
            CheckFontCompatibility = true
        };

        // Save the workbook as a PDF with all fonts embedded
        workbook.Save("EmbeddedFonts.pdf", pdfOptions);
    }
}
