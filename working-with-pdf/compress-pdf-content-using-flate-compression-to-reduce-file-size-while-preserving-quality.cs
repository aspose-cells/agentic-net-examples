using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfCompressionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Sample PDF Compression Demo");
            sheet.Cells["A2"].PutValue("This PDF is saved using Flate compression.");

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set the core compression algorithm to Flate (compresses all content except images)
            pdfOptions.PdfCompression = PdfCompressionCore.Flate;

            // Optional: prioritize smaller file size over print quality
            pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

            // Save the workbook as a PDF with the specified compression settings
            workbook.Save("CompressedOutput.pdf", pdfOptions);

            Console.WriteLine("PDF saved with Flate compression.");
        }
    }
}