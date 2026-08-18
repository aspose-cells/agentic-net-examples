// Title: C# – Convert Aspose.Cells Workbook to PDF with lossless Flate compression
// Description: Creates a workbook, fills it with sample data, configures PdfSaveOptions to use PdfCompressionCore.Flate, and saves the file as a PDF. The Flate compression reduces size while keeping image and chart quality intact.
// Keywords: Aspose.Cells PDF export C# | lossless PDF compression Aspose | PdfCompressionCore.Flate example | high‑quality Excel to PDF conversion | .NET workbook to PDF | Flate compression for PDFs
// Common Searches: Aspose.Cells save workbook as PDF with lossless compression | C# set PdfCompressionCore.Flate when exporting Excel to PDF | how to keep image quality in Aspose.Cells PDF output | reduce PDF size without losing fidelity using Aspose.Cells
// Developer Intent: Export an Excel workbook to PDF while applying Flate compression to retain visual fidelity.
// Use Cases: Generate financial statements as PDFs that preserve chart clarity. | Create printable invoices with high‑resolution graphics and modest file size. | Archive regulatory spreadsheets as PDFs without degrading images or layout.
// AI Prompts: Provide C# code that converts an Aspose.Cells workbook to PDF using PdfCompressionCore.Flate. | Show how to configure PdfSaveOptions for lossless compression in Aspose.Cells .NET. | Explain the impact of Flate compression on PDF size and image quality when exporting from Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfConversion
{
    // Creates a workbook, fills it with sample data, configures PdfSaveOptions to use PdfCompressionCore.Flate, and saves the file as a PDF. The Flate compression reduces size while keeping image and chart quality intact.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Apply lossless compression (Flate) to all PDF content except images
            // This preserves visual fidelity while reducing file size
            pdfOptions.PdfCompression = PdfCompressionCore.Flate;

            // Optionally, keep the default optimization (high quality)
            // pdfOptions.OptimizationType = PdfOptimizationType.Standard;

            // Save the workbook as a PDF using the configured options
            string outputPath = "Workbook_Lossless.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved to PDF with lossless compression: {outputPath}");
        }
    }
}
