using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class EmbedFontsPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample text
        worksheet.Cells["A1"].PutValue("Sample text with a specific font");

        // Apply a font that may need to be embedded
        Style style = worksheet.Cells["A1"].GetStyle();
        style.Font.Name = "Times New Roman"; // choose a font that is not guaranteed to be present on all devices
        style.Font.Size = 14;
        worksheet.Cells["A1"].SetStyle(style);

        // Configure PDF save options to embed all fonts
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            // Use Identity encoding to ensure full Unicode support
            FontEncoding = PdfFontEncoding.Identity,

            // Embed TrueType fonts (standard Windows fonts are embedded by default)
            EmbedStandardWindowsFonts = true,

            // Set a default font in case a cell does not specify one
            DefaultFont = "Times New Roman"
        };

        // Save the workbook as a PDF with embedded fonts
        workbook.Save("EmbeddedFonts.pdf", pdfSaveOptions);
    }
}