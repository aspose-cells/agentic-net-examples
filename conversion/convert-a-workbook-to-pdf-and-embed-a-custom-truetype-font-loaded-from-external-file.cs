using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class EmbedCustomFontPdf
{
    static void Main()
    {
        // Path to the folder that contains the custom TrueType font file.
        // Set to true to scan subfolders as well.
        FontConfigs.SetFontFolder(@"C:\CustomFonts", true);

        // Create a new workbook and add some sample text.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello with Custom Font");

        // Apply a style that uses the custom font.
        // The font name should match the name defined inside the TTF file (without extension).
        Style style = sheet.Cells["A1"].GetStyle();
        style.Font.Name = "MyCustomFont";
        sheet.Cells["A1"].SetStyle(style);

        // Configure PDF save options to embed the custom font.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use the custom font as the default when rendering the PDF.
            DefaultFont = "MyCustomFont",

            // Try to use the workbook's default font first (helps with Unicode characters).
            CheckWorkbookDefaultFont = true,

            // Use Identity encoding for all embedded fonts.
            FontEncoding = PdfFontEncoding.Identity,

            // Ensure standard Windows fonts are embedded (does not affect custom fonts but keeps consistency).
            EmbedStandardWindowsFonts = true
        };

        // Save the workbook as a PDF with the specified options.
        workbook.Save("output.pdf", pdfOptions);
    }
}