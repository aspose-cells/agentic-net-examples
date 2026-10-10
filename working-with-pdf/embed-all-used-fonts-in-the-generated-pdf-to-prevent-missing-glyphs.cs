// Title: How to embed every used font when converting an Aspose.Cells workbook to PDF in C#
// AI Prompts: Create C# code that converts an Aspose.Cells workbook to PDF and forces embedding of all referenced fonts. | Show how to configure PdfSaveOptions to embed Windows system fonts during PDF export. | Explain a fallback method for font embedding when the dedicated FontEmbeddingMode API is unavailable in Aspose.Cells.
// Common Searches: Aspose.Cells C# export to PDF with embedded fonts for Chinese characters | how to ensure fonts are embedded in PDF generated from Excel using Aspose.Cells | example of PdfSaveOptions to include Windows fonts in PDF via C# | avoid glyph loss in PDF generated from an Aspose.Cells workbook | workaround for font embedding when Aspose.Cells lacks direct embedding mode
// Tags: Aspose.Cells PDF font embedding C# | PdfSaveOptions Windows system font embedding | embed non-Latin fonts in PDF export | prevent missing glyphs in Aspose.Cells PDF | Excel to PDF font inclusion settings

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds Unicode text, applies a font that supports the characters, configures PdfSaveOptions with EmbedStandardWindowsFonts enabled, and saves the workbook as a PDF, ensuring all used fonts are embedded to avoid missing glyphs.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some text that uses non‑Latin glyphs
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Hello, 世界!"); // Sample text with Chinese characters

            // Apply a font that supports the characters
            Style style = workbook.CreateStyle();
            style.Font.Name = "Arial Unicode MS";
            style.Font.Size = 14;
            cell.SetStyle(style);

            // Configure PDF save options to embed standard Windows fonts
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure standard Windows fonts are embedded (fallback safety)
                EmbedStandardWindowsFonts = true
                // Note: FontEmbeddingMode property is not available in this version of Aspose.Cells
            };

            // Define output path
            string outputPath = "EmbeddedFontsOutput.pdf";

            // Save the workbook as a PDF with the embedding options applied
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
