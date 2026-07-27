// Title: Convert Excel to PDF with lossless image compression using Aspose.Cells for .NET (C#)
// Description: Loads an XLSX workbook, enables automatic lossless picture compression, applies Flate compression and minimum‑size PDF optimization, then saves the file as a compact PDF while preserving image quality.
// Keywords: Aspose.Cells PDF conversion | C# Excel to PDF | lossless image compression | AutoCompressPictures | Flate PDF compression | minimum size PDF optimization | .NET workbook to PDF | compress embedded pictures
// Common Searches: Aspose.Cells enable lossless picture compression | C# convert Excel to PDF with small file size | how to use PdfSaveOptions for minimum PDF size | auto compress pictures when saving PDF in Aspose.Cells | Flate compression for PDF generated from Excel
// Developer Intent: Generate a PDF from an Excel workbook in .NET while applying lossless compression to embedded images to keep the file size low without degrading visual quality.
// Use Cases: Produce PDF reports from financial spreadsheets where charts and photos must stay crisp but the document size should be minimal. | Automate batch conversion of dozens of .xlsx files to archived PDFs with lossless image handling. | Create PDF invoices from Excel templates that contain logos, ensuring the logos are compressed without quality loss.
// AI Prompts: Show how to set a custom DPI for the PDF output while keeping AutoCompressPictures enabled. | Provide a loop that converts a list of Excel files to PDFs, applying lossless image compression to each. | Explain how to programmatically verify that images in the generated PDF are losslessly compressed using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an XLSX workbook, enables automatic lossless picture compression, applies Flate compression and minimum‑size PDF optimization, then saves the file as a compact PDF while preserving image quality.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Enable automatic lossless compression of embedded pictures
        workbook.Settings.AutoCompressPictures = true;

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Use Flate compression for all PDF content except images (lossless)
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;

        // Optimize for minimum file size while keeping quality
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
