// Title: Compress PDF with Flate and high‑resolution image resampling using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, configures PdfSaveOptions to apply Flate compression, selects the MinimumSize optimization mode, resamples images to 220 DPI at 90 % JPEG quality, and saves the result as a compact PDF while preserving visual fidelity.
// Keywords: Aspose.Cells | PdfSaveOptions | PDF compression | Flate compression | image resampling | C# | .NET | minimum size PDF | high DPI images | PDF optimization | Aspose.Cells PDF export
// Common Searches: Aspose.Cells PDF compression Flate C# | How to reduce PDF size with Aspose.Cells | Set image resample DPI in Aspose.Cells PDF | MinimumSize PDF optimization Aspose.Cells | Custom PdfSaveOptions example .NET
// Developer Intent: Generate a PDF from an Excel workbook with tailored compression settings that shrink file size without sacrificing image quality.
// Use Cases: Email‑ready reports that contain high‑resolution charts but must stay under attachment limits. | Archival documents where storage cost is critical yet images need to remain clear. | Web‑downloadable PDFs for bandwidth‑constrained users while keeping visual detail.
// AI Prompts: Show how to switch PdfSaveOptions to LZW compression in Aspose.Cells. | Provide C# code to add a watermark to a PDF while using custom compression. | Explain how to balance DPI and JPEG quality for different PDF size‑quality requirements with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook, configures PdfSaveOptions to apply Flate compression, selects the MinimumSize optimization mode, resamples images to 220 DPI at 90 % JPEG quality, and saves the result as a compact PDF while preserving visual fidelity.
class PdfCompressionDemo
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF Compression Demo");
        sheet.Cells["A2"].PutValue("Demonstrates custom compression while preserving image quality");

        // Optionally add a picture to illustrate image handling (ensure the file exists)
        // sheet.Pictures.Add(5, 0, "sample.jpg");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Use Flate compression for the PDF core content (good compression ratio)
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;

        // Optimize for minimum file size while keeping acceptable quality
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Resample images to a high PPI (e.g., 220) with high JPEG quality (90%)
        // This maintains image quality but still benefits from compression
        pdfOptions.SetImageResample(220, 90);

        // Save the workbook as a PDF using the configured options
        workbook.Save("CompressedOutput.pdf", pdfOptions);
    }
}
