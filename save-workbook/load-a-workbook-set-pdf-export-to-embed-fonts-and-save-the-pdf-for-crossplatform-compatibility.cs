// Title: C# – Save Excel as PDF with Embedded Fonts using Aspose.Cells
// Description: Load an Excel workbook with Aspose.Cells, set PdfSaveOptions to embed TrueType fonts, apply Identity‑H encoding and default‑font fallback, then save as a PDF that renders consistently on Windows, macOS and Linux.
// Keywords: Aspose.Cells | C# | PdfSaveOptions | EmbedStandardWindowsFonts | Identity-H | font embedding | Excel to PDF | cross‑platform PDF | PDF export options | preserve fonts
// Common Searches: Aspose.Cells embed fonts PDF | PdfSaveOptions embed standard windows fonts C# | How to keep fonts when converting Excel to PDF Aspose | Identity-H encoding Aspose.Cells PDF | Cross‑platform PDF from Excel .NET | Save workbook as PDF with font embedding
// Developer Intent: Create a PDF from an Excel workbook with all fonts embedded so the document looks identical on any platform.
// Use Cases: Automated generation of PDF reports from Excel templates that retain exact typography on all operating systems. | Production of printable invoices or contracts where missing fonts could cause layout issues for recipients. | Archiving financial statements as PDFs that meet compliance requirements for long‑term readability.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells and saves it as a PDF with embedded TrueType fonts and Identity‑H encoding. | Explain the effect of each PdfSaveOptions property (EmbedStandardWindowsFonts, FontEncoding, CheckWorkbookDefaultFont) on font handling in the exported PDF. | Provide a step‑by‑step tutorial for configuring Aspose.Cells PDF export to achieve cross‑platform font compatibility. | Troubleshoot why certain custom fonts are not appearing in the PDF after using PdfSaveOptions with font embedding.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an Excel workbook with Aspose.Cells, set PdfSaveOptions to embed TrueType fonts, apply Identity‑H encoding and default‑font fallback, then save as a PDF that renders consistently on Windows, macOS and Linux.
class Program
{
    static void Main()
    {
        // Load an existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to embed fonts for cross‑platform compatibility
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.EmbedStandardWindowsFonts = true;          // Embed TrueType fonts
        pdfOptions.FontEncoding = PdfFontEncoding.Identity; // Use Identity-H encoding for all fonts
        pdfOptions.CheckWorkbookDefaultFont = true;          // Use workbook's default font when needed

        // Save the workbook as a PDF file with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
