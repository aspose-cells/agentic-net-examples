using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfOptimizationExample
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Aspose.Cells PDF optimization with MinimumSize and font subsetting.");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // 1. Optimize for minimum file size (print quality is less important)
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // 2. Enable font subsetting by using ANSI preferred encoding.
        //    This reduces the embedded font data to only the characters used.
        pdfOptions.FontEncoding = PdfFontEncoding.AnsiPrefer;

        // Optional: ensure standard Windows fonts are embedded (default is true)
        pdfOptions.EmbedStandardWindowsFonts = true;

        // Save the workbook as a PDF with the specified options
        workbook.Save("OptimizedWithFontSubsetting.pdf", pdfOptions);
    }
}