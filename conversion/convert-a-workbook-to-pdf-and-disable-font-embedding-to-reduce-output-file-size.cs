using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample text for PDF conversion.");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Disable embedding of standard Windows fonts to reduce file size
        pdfOptions.EmbedStandardWindowsFonts = false;

        // Optional: set optimization type to prioritize minimum file size
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}