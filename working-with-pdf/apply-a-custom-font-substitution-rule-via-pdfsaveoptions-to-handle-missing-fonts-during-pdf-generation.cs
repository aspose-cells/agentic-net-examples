// Title: Custom Font Substitution with PdfSaveOptions in Aspose.Cells for .NET
// Description: Shows how to set font substitutes for missing fonts, configure PdfSaveOptions (DefaultFont, CheckWorkbookDefaultFont, IsFontSubstitutionCharGranularity, CheckFontCompatibility) and export a workbook to PDF using Aspose.Cells C#.
// Keywords: Aspose.Cells | PdfSaveOptions | font substitution | SetFontSubstitutes | missing fonts | default font | character granularity | font compatibility | C# | .NET | PDF export
// Common Searches: Aspose.Cells font substitution PDF | PdfSaveOptions default font .NET | SetFontSubstitutes example C# | handle missing fonts when saving Excel to PDF | character level font fallback Aspose.Cells
// Developer Intent: Configure Aspose.Cells to replace unavailable fonts with defined substitutes and generate a PDF with reliable fallback handling.
// Use Cases: Replace a non‑installed font (e.g., NonExistentFont) with Arial or Helvetica during PDF conversion. | Apply character‑granular substitution to correctly render mixed‑script text in cells. | Set a default fallback font and enable compatibility checks to avoid missing‑glyph errors in the output PDF.
// AI Prompts: Write C# code that registers a font substitution list for a missing font and saves an Aspose.Cells workbook as PDF using PdfSaveOptions with appropriate fallback settings. | Explain the effect of each PdfSaveOptions property—DefaultFont, CheckWorkbookDefaultFont, IsFontSubstitutionCharGranularity, CheckFontCompatibility—on font fallback during PDF export.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to set font substitutes for missing fonts, configure PdfSaveOptions (DefaultFont, CheckWorkbookDefaultFont, IsFontSubstitutionCharGranularity, CheckFontCompatibility) and export a workbook to PDF using Aspose.Cells C#.
class FontSubstitutionPdfDemo
{
    static void Main()
    {
        // Define substitutes for a font that might be missing on the system.
        // If "NonExistentFont" cannot be found, Aspose.Cells will try "Arial" then "Helvetica".
        FontConfigs.SetFontSubstitutes("NonExistentFont", new string[] { "Arial", "Helvetica" });

        // Create a new workbook and write some text using the missing font.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Text with a font that may be missing");

        // Apply the missing font to the cell.
        Style style = workbook.CreateStyle();
        style.Font.Name = "NonExistentFont";
        style.Font.Size = 14;
        cell.SetStyle(style);

        // Configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use a known font as the default when no suitable font is found.
            DefaultFont = "Arial",
            // Try the workbook's default font first before falling back to system fonts.
            CheckWorkbookDefaultFont = true,
            // Substitute fonts at character granularity for better rendering of mixed scripts.
            IsFontSubstitutionCharGranularity = true,
            // Ensure compatibility checking is enabled so fallback fonts are searched.
            CheckFontCompatibility = true
        };

        // Save the workbook as a PDF using the configured options.
        workbook.Save("FontSubstitutionDemo.pdf", pdfOptions);
    }
}
