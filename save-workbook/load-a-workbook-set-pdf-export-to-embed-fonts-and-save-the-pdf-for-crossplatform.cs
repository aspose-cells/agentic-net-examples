using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to embed fonts for cross‑platform compatibility
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Embed TrueType fonts (affects ASCII characters 32‑127; non‑ASCII are always embedded)
        pdfOptions.EmbedStandardWindowsFonts = true;

        // Use Identity encoding to preserve all characters in the PDF
        pdfOptions.FontEncoding = PdfFontEncoding.Identity;

        // Try to use the workbook's default font for Unicode characters
        pdfOptions.CheckWorkbookDefaultFont = true;

        // Set a widely available default font as a fallback
        pdfOptions.DefaultFont = "Arial";

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}