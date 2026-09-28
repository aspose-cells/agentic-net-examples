// Title: How to convert an Excel workbook to PDF while applying lossless image compression with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file, enables the workbook's AutoCompressPictures setting, configures PdfSaveOptions to use Flate compression and MinimumSize optimization, and saves the workbook as a PDF. | Show how to set PdfCompressionCore.Flate and PdfOptimizationType.MinimumSize in Aspose.Cells to produce a PDF whose embedded images are losslessly compressed.
// Common Searches: C# Aspose.Cells how to enable lossless picture compression when saving Excel as PDF | Save workbook to PDF with smallest file size using Flate compression in Aspose.Cells | Effect of AutoCompressPictures property on PDF image size in Aspose.Cells | PdfSaveOptions example for minimum size PDF in .NET | Compress embedded images in PDF generated from .xlsx using Aspose.Cells
// Tags: Aspose.Cells AutoCompressPictures setting | PdfSaveOptions Flate compression .NET | Workbook to PDF minimum size optimization | C# lossless PDF image compression | Excel to PDF embedded image compression Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads an Excel workbook, activates automatic lossless picture compression, configures PDF save options with Flate compression and minimum‑size optimization, and saves the output as a compressed PDF.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Enable automatic lossless compression of pictures inside the workbook
        workbook.Settings.AutoCompressPictures = true;

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Use a lossless compression algorithm for PDF content (Flate = ZIP)
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;

        // Optimize the PDF for minimum file size (helps compress embedded images)
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
