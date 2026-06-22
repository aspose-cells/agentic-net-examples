using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;   // Required for PdfFontEncoding enum

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data (including Unicode characters to demonstrate font embedding)
        worksheet.Cells["A1"].PutValue("Sample text with Unicode 漢字");

        // Configure PDF save options to embed fonts and preserve text appearance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure TrueType fonts are embedded (default is true, set explicitly for clarity)
            EmbedStandardWindowsFonts = true,

            // Use Identity encoding for all embedded fonts
            FontEncoding = PdfFontEncoding.Identity,

            // Specify a fallback default font (e.g., Arial) in case a cell's font is missing
            DefaultFont = "Arial",

            // Try to use the workbook's default font for Unicode characters
            CheckWorkbookDefaultFont = true
        };

        // Save the worksheet as a PDF file with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}