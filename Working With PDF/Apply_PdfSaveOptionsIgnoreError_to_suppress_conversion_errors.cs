using System;
using Aspose.Cells;

namespace AsposeCellsIgnoreErrorDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data that might cause rendering issues
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Sample Data");
            // Example of a potential problematic element (commented out):
            // worksheet.Pictures.Add(0, 0, "unsupported_image.bmp");

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Enable suppression of rendering errors
            pdfOptions.IgnoreError = true;

            // Save the workbook as PDF using the configured options
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}