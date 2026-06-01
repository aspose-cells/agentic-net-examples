using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample text with embedded fonts");
        sheet.Cells["A2"].PutValue("中文字符测试");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set PDF/A‑1a compliance
        pdfOptions.Compliance = PdfCompliance.PdfA1a;

        // Ensure all fonts are embedded (default is true, set explicitly for clarity)
        pdfOptions.EmbedStandardWindowsFonts = true;

        // Provide a default font for Unicode characters
        pdfOptions.DefaultFont = "Arial";

        // Use Identity encoding for all embedded fonts
        pdfOptions.FontEncoding = PdfFontEncoding.Identity;

        // Save the workbook as a PDF file with the specified options
        workbook.Save("Output_PdfA1a.pdf", pdfOptions);
    }
}