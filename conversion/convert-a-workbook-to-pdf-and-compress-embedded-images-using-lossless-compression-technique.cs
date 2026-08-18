// Title: C# – Convert Excel to PDF with lossless image compression using Aspose.Cells
// Description: Load an Excel workbook, enable automatic picture compression, set PdfSaveOptions to Flate compression and MinimumSize optimization, and save the file as a PDF where embedded images are compressed without quality loss.
// Keywords: Aspose.Cells | C# PDF conversion | lossless image compression | AutoCompressPictures | PdfSaveOptions | Flate compression | MinimumSize optimization | Excel to PDF | embedded images | file size reduction
// Common Searches: Aspose.Cells convert xlsx to pdf lossless | C# enable auto picture compression Aspose.Cells | PdfSaveOptions Flate compression example | How to reduce PDF size with Aspose.Cells | MinimumSize optimization Aspose.Cells PDF export
// Developer Intent: Create a PDF from an Excel workbook while preserving image quality through lossless compression.
// Use Cases: Generate high‑fidelity PDF reports from financial spreadsheets that contain charts and photos. | Archive Excel workbooks as compact PDFs for document management systems without degrading image clarity. | Provide downloadable PDFs on a web portal where image detail must remain intact while keeping file size low.
// AI Prompts: Write C# code that converts an .xlsx file to PDF using Aspose.Cells with lossless image compression. | Explain the effect of PdfCompression.Flate and PdfOptimizationType.MinimumSize on embedded images in a PDF. | Provide a step‑by‑step tutorial for configuring PdfSaveOptions to achieve lossless compression of pictures in an Excel‑to‑PDF conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfCompressionDemo
{
    // Load an Excel workbook, enable automatic picture compression, set PdfSaveOptions to Flate compression and MinimumSize optimization, and save the file as a PDF where embedded images are compressed without quality loss.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your source file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Optional: enable automatic picture compression inside the workbook
            workbook.Settings.AutoCompressPictures = true;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Use lossless Flate compression for PDF content (excluding images)
            pdfOptions.PdfCompression = PdfCompressionCore.Flate;

            // Optimize for minimum file size (helps compress embedded images losslessly)
            pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

            // Save the workbook as a PDF with the specified compression settings
            workbook.Save("output_compressed.pdf", pdfOptions);
        }
    }
}
