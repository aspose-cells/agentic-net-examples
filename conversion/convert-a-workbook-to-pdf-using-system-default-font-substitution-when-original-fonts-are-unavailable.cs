// Title: C# – Convert Excel to PDF with System Font Substitution using Aspose.Cells
// Description: Loads an Excel workbook, enables FontConfigs.PreferSystemFontSubstitutes, sets PdfSaveOptions.CheckWorkbookDefaultFont, and saves the file as PDF without a custom DefaultFont. Missing fonts are automatically replaced by the operating system’s fonts, ensuring consistent output on any server.
// Keywords: Aspose.Cells PDF conversion | C# Excel to PDF | system font substitution | FontConfigs.PreferSystemFontSubstitutes | PdfSaveOptions.CheckWorkbookDefaultFont | default font fallback | missing font handling | server‑side PDF generation | global font compatibility
// Common Searches: Aspose.Cells enable system font substitutes | C# convert Excel to PDF without specifying DefaultFont | PdfSaveOptions CheckWorkbookDefaultFont example | how to fallback to OS fonts in Aspose.Cells PDF export | Excel to PDF font substitution C#
// Developer Intent: Create a PDF from an Excel workbook while automatically falling back to the OS’s fonts for any unavailable typefaces.
// Use Cases: Generate printable reports on a server that lacks custom corporate fonts. | Batch‑process Excel files into PDFs in a cloud environment with limited font resources. | Produce invoices or certificates from templates that reference fonts not installed on the deployment machine.
// AI Prompts: Show C# code that converts an Excel file to PDF with Aspose.Cells using system font substitution. | Explain how FontConfigs.PreferSystemFontSubstitutes and PdfSaveOptions.CheckWorkbookDefaultFont work together. | Give troubleshooting steps for missing font warnings when exporting Excel to PDF with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an Excel workbook, enables FontConfigs.PreferSystemFontSubstitutes, sets PdfSaveOptions.CheckWorkbookDefaultFont, and saves the file as PDF without a custom DefaultFont. Missing fonts are automatically replaced by the operating system’s fonts, ensuring consistent output on any server.
class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Use system font substitutes when the original font is not available
        FontConfigs.PreferSystemFontSubstitutes = true;

        // Load the source workbook (replace the path with your actual file)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Ensure the workbook's default font is checked (default is true)
        pdfOptions.CheckWorkbookDefaultFont = true;

        // Do not set DefaultFont so Aspose.Cells will fall back to the system default font
        // pdfOptions.DefaultFont = null; // optional, shown for clarity

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
