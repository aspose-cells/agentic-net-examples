using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class CustomPaperSizePdfDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data to demonstrate the output
        sheet.Cells["A1"].PutValue("Custom Paper Size with 300 DPI PDF");
        sheet.Cells["A2"].PutValue("This PDF should have higher image quality.");

        // Set a custom paper size (width: 4 inches, height: 6 inches)
        sheet.PageSetup.CustomPaperSize(4.0, 6.0);

        // Optionally set the print quality to match the desired DPI
        sheet.PageSetup.PrintQuality = 300;

        // Configure PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Resample images to 300 PPI and use high JPEG quality (90%)
        pdfSaveOptions.SetImageResample(300, 90);

        // Use standard optimization for high print quality
        pdfSaveOptions.OptimizationType = PdfOptimizationType.Standard;

        // Save the workbook as a PDF with the specified options
        workbook.Save("CustomPaper_300DPI.pdf", pdfSaveOptions);
    }
}