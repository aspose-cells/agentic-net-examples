using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfExportWithEmbeddedFonts
{
    static void Main()
    {
        // Path to the source Excel workbook
        string excelPath = "input.xlsx";

        // Path for the resulting PDF file
        string pdfPath = "output.pdf";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(excelPath);

        // Configure PDF save options to embed fonts and ensure cross‑platform rendering
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed TrueType fonts (ASCII range) – required for proper display on other platforms
            EmbedStandardWindowsFonts = true,

            // Use Identity encoding for all embedded fonts (covers Unicode characters)
            FontEncoding = PdfFontEncoding.Identity,

            // Try to use the workbook's default font when a cell's font is missing or unsupported
            CheckWorkbookDefaultFont = true,

            // Specify a fallback font (e.g., Arial) for Unicode characters not covered by cell styles
            DefaultFont = "Arial"
        };

        // Save the workbook as a PDF using the configured options
        workbook.Save(pdfPath, pdfOptions);

        Console.WriteLine("PDF saved with embedded fonts at: " + pdfPath);
    }
}