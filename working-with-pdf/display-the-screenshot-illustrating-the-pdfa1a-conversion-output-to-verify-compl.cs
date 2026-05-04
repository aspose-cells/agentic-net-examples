using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class PdfA1aComplianceDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("PDF/A-1a compliance demonstration");
        worksheet.Cells["A2"].PutValue(DateTime.Now);
        worksheet.Cells["A3"].PutValue(12345);

        // Configure PDF save options to use PDF/A-1a compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a
        };

        // Save the workbook as a PDF file with the specified compliance level
        string pdfPath = "PdfA1aOutput.pdf";
        workbook.Save(pdfPath, pdfOptions);
        Console.WriteLine($"PDF saved with PDF/A-1a compliance: {pdfPath}");

        // Render the first page of the worksheet to an image for visual verification
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png
        };
        SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
        string imagePath = "PdfA1aScreenshot.png";
        sheetRender.ToImage(0, imagePath);
        Console.WriteLine($"Worksheet rendered to image: {imagePath}");
    }
}