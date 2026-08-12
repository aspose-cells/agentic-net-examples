// Title: Compress Excel‑to‑PDF with Flate compression using Aspose.Cells for .NET
// Description: Shows how to create a workbook, set PdfSaveOptions.PdfCompression to PdfCompressionCore.Flate, choose MinimumSize optimization, and save the result as a reduced‑size PDF (CompressedOutput.pdf) with C#.
// Keywords: Aspose.Cells PDF compression | Flate compression .NET | PdfSaveOptions | PdfCompressionCore.Flate | minimum PDF size | C# Aspose.Cells PDF | reduce PDF file size | Excel to PDF compression | Aspose.Cells optimization | PDF output size reduction
// Common Searches: Aspose.Cells Flate compression C# | How to reduce PDF size with Aspose.Cells | PdfSaveOptions MinimumSize example | Compress Excel PDF using Aspose.Cells .NET | Set PDF compression type to Flate Aspose | Aspose.Cells PDF optimization for small files
// Developer Intent: Generate a PDF from an Excel workbook with Flate compression to achieve the smallest possible file size while preserving text quality.
// Use Cases: Email large financial reports as compact PDFs to stay under attachment limits. | Archive spreadsheets as low‑size PDFs to cut cloud storage costs. | Serve downloadable PDFs on bandwidth‑constrained websites for faster load times. | Automate batch conversion of Excel files to minimal‑size PDFs in CI/CD pipelines.
// AI Prompts: Write C# code that saves a Workbook as a PDF using PdfCompressionCore.Flate and MinimumSize optimization. | Explain the differences between Flate, JPEG, and CCITT compression options in Aspose.Cells PDF export. | Show how to apply Flate compression to all PDF streams except images with Aspose.Cells. | Provide a step‑by‑step guide to reduce PDF size when converting Excel to PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfCompressionDemo
{
    // Shows how to create a workbook, set PdfSaveOptions.PdfCompression to PdfCompressionCore.Flate, choose MinimumSize optimization, and save the result as a reduced‑size PDF (CompressedOutput.pdf) with C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("PDF Compression Demo");
            sheet.Cells["A2"].PutValue("This PDF is saved using Flate compression to reduce file size.");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Use Flate compression for all PDF content except images
            pdfOptions.PdfCompression = PdfCompressionCore.Flate;

            // Optional: prioritize smaller file size over print quality
            pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

            // Save the workbook as a PDF with the specified compression settings
            workbook.Save("CompressedOutput.pdf", pdfOptions);

            Console.WriteLine("PDF saved with Flate compression.");
        }
    }
}
