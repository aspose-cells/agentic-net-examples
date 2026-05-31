using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class EmbedAllFontsPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data
        worksheet.Cells["A1"].PutValue("Text with default font");
        worksheet.Cells["A2"].PutValue("Text with custom font");

        // Apply a custom font style to demonstrate embedding of a non‑default font
        Style customStyle = workbook.CreateStyle();
        customStyle.Font.Name = "Times New Roman";
        customStyle.Font.Size = 12;
        worksheet.Cells["A2"].SetStyle(customStyle);

        // If you have additional fonts in a custom folder, uncomment the line below
        // FontConfigs.SetFontFolder(@"C:\MyFonts", true);

        // Configure PDF save options to embed all used fonts
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use Identity-H encoding to embed fonts fully
            FontEncoding = PdfFontEncoding.Identity,
            // Ensure standard Windows TrueType fonts are embedded
            EmbedStandardWindowsFonts = true,
            // Provide a fallback default font
            DefaultFont = "Arial"
        };

        // Save the workbook as a PDF with the specified options
        workbook.Save("EmbeddedFontsOutput.pdf", pdfOptions);
    }
}