// Title: Compress PDF output with Flate compression using Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add sample data, configure PdfSaveOptions.PdfCompression to PdfCompressionCore.Flate, and save a smaller, lossless PDF file.
// Keywords: Aspose.Cells | C# | .NET | Flate compression | PDF size reduction | PdfSaveOptions | lossless PDF | Excel to PDF | compress PDF | PdfCompressionCore.Flate
// Common Searches: Aspose.Cells Flate PDF compression C# | How to reduce PDF size with Aspose.Cells .NET | PdfSaveOptions Flate example | Compress Excel workbook to PDF lossless | Set PDF compression type Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook using lossless Flate compression to minimize file size without sacrificing quality.
// Use Cases: Emailing Excel‑derived reports where attachment size matters. | Batch exporting workbooks to archival PDFs with consistent, lossless compression. | Producing high‑quality printable PDFs for client deliverables while keeping file size low.
// AI Prompts: Provide a C# snippet that sets additional PDF metadata (author, title) while using Flate compression with Aspose.Cells. | Show how to loop through multiple workbooks, apply PdfCompressionCore.Flate, and handle errors during PDF conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to create a workbook, add sample data, configure PdfSaveOptions.PdfCompression to PdfCompressionCore.Flate, and save a smaller, lossless PDF file.
class PdfCompressionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Populate the first worksheet with sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF Compression Demo");
        sheet.Cells["A2"].PutValue("This PDF is saved using Flate compression to reduce size while keeping quality.");

        // Set up PDF save options to use Flate compression (lossless)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;

        // Save the workbook as a PDF file with the specified compression
        workbook.Save("CompressedOutput.pdf", pdfOptions);
    }
}
