// Title: Export an Excel workbook to PDF with embedded fonts using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, configures PdfSaveOptions to embed standard Windows TrueType fonts, sets Identity font encoding and a fallback font, then saves the workbook as a PDF. | Demonstrate how to enable font embedding and compatibility checks in Aspose.Cells when converting Excel to PDF for reliable rendering on any platform.
// Common Searches: Aspose.Cells C# embed fonts in PDF export | How to set default font for PDF conversion with Aspose.Cells | PdfSaveOptions FontEncoding Identity Aspose.Cells example | Enable font compatibility checking when saving Excel as PDF using Aspose.Cells | Generate cross‑platform PDF from Excel with embedded TrueType fonts in .NET
// Tags: Aspose.Cells PdfSaveOptions embed fonts | C# Aspose.Cells export Excel to PDF | Identity font encoding Aspose.Cells PDF | default font fallback Aspose.Cells PDF | font compatibility check Aspose.Cells PDF

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The code loads 'input.xlsx' into an Aspose.Cells Workbook, configures PdfSaveOptions to embed standard Windows TrueType fonts, uses Identity encoding, sets Arial as a default fallback font, enables workbook default font checking and font compatibility verification, and then saves the workbook as 'output.pdf' with all fonts embedded for consistent cross‑platform rendering.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to embed fonts for cross‑platform compatibility
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Ensure standard Windows TrueType fonts are embedded (default is true, set explicitly)
        pdfOptions.EmbedStandardWindowsFonts = true;

        // Use Identity encoding for all embedded fonts
        pdfOptions.FontEncoding = PdfFontEncoding.Identity;

        // Specify a common default font in case cell styles lack proper font information
        pdfOptions.DefaultFont = "Arial";

        // Attempt to use the workbook's default font first for Unicode characters
        pdfOptions.CheckWorkbookDefaultFont = true;

        // Keep font compatibility checking enabled for fallback substitution
        pdfOptions.CheckFontCompatibility = true;

        // Save the workbook as a PDF file with the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
