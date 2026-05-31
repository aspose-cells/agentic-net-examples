using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlToPdf
{
    class Program
    {
        static void Main()
        {
            // Load the HTML file into a workbook
            Workbook workbook = new Workbook("input.html");

            // Configure PDF save options with object stream compression (Flate)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Compress all non‑image content using Flate algorithm
                PdfCompression = PdfCompressionCore.Flate,
                // Optional: minimize file size
                OptimizationType = PdfOptimizationType.MinimumSize
            };

            // Save the workbook as a compressed PDF
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}