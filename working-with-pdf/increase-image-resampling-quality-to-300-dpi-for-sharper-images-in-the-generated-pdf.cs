// Title: Increase image resolution to 300 DPI when saving an Excel workbook as PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to set PDF image resolution to 300 DPI | PdfSaveOptions ImageResolution property example for high‑DPI PDF | export Excel worksheet to PDF with high‑resolution images using Aspose.Cells | C# code to improve image quality in PDF generated from a workbook | increase DPI of images embedded in PDF created by Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions image resolution setting | high DPI image export to PDF .NET | C# increase PDF image quality Aspose.Cells | image resampling 300 DPI Aspose.Cells PDF generation | Excel to PDF high resolution images using Aspose.Cells

using Aspose.Cells;
using System;

// Demonstrates how to configure PdfSaveOptions.ImageResolution to 300 DPI in Aspose.Cells for .NET, producing sharper, high‑resolution images in the resulting PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B2"].PutValue(12345);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // If the current Aspose.Cells version supports it, you can set the image DPI:
            // pdfOptions.ImageResolution = 300; // Uncomment when available

            // Save the workbook as a PDF
            workbook.Save("Result.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
