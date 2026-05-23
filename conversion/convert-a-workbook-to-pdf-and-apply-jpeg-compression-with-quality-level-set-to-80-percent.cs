using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions related enums if needed

class Program
{
    static void Main()
    {
        // Create a new workbook (you can also load an existing file with new Workbook("input.xlsx"))
        Workbook workbook = new Workbook();

        // Add some sample data to demonstrate the PDF content
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Demo: PDF with JPEG compression (quality 80%)");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set image resampling: desired PPI (e.g., 96) and JPEG quality (80%)
        // This will compress all images in the PDF using JPEG with the specified quality.
        pdfOptions.SetImageResample(96, 80);

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}