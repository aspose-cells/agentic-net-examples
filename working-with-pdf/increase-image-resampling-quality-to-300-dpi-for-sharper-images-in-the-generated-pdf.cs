using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to illustrate the PDF content
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for PDF with high‑DPI images");
        sheet.Cells["A2"].PutValue("Additional content");

        // Configure PDF save options to resample images at 300 DPI with maximum JPEG quality
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.SetImageResample(300, 100); // desiredPPI = 300, jpegQuality = 100%

        // Save the workbook as a PDF using the configured options
        workbook.Save("HighDPI_Output.pdf", pdfOptions);
    }
}