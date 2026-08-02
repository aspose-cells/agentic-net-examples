using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Configure PDF save options to embed fonts for cross‑platform compatibility
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure TrueType fonts are embedded (required for ASCII characters 32‑127)
            pdfOptions.EmbedStandardWindowsFonts = true;

            // Use Identity encoding so all embedded fonts are correctly referenced
            pdfOptions.FontEncoding = PdfFontEncoding.Identity;

            // Optionally set a default font to handle Unicode characters without explicit font styling
            pdfOptions.DefaultFont = "Arial";

            // Save the workbook as a PDF using the configured options
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}