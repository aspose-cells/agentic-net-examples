using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class FontSubstitutionPdfDemo
{
    static void Main()
    {
        // Define substitute fonts for a font that may be missing on the system.
        // If "NonExistentFont" is not found, Aspose.Cells will try "Arial" then "Helvetica".
        FontConfigs.SetFontSubstitutes("NonExistentFont", new string[] { "Arial", "Helvetica" });

        // Create a new workbook and add a cell that uses the missing font.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Text using a missing font");

        // Apply the missing font to the cell style.
        Style style = workbook.CreateStyle();
        style.Font.Name = "NonExistentFont";
        sheet.Cells["A1"].SetStyle(style);

        // Configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use a known default font when the original font cannot be rendered.
            DefaultFont = "Arial",
            // Try the workbook's default font first before falling back to system fonts.
            CheckWorkbookDefaultFont = true,
            // Substitute fonts at character granularity to avoid whole-cell substitution.
            IsFontSubstitutionCharGranularity = true,
            // Use Identity encoding for embedded fonts.
            FontEncoding = PdfFontEncoding.Identity
        };

        // Save the workbook as a PDF using the configured options.
        workbook.Save("FontSubstitutionOutput.pdf", pdfOptions);
    }
}