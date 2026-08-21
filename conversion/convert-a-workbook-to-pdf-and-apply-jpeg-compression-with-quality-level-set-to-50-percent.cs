// Title: Convert Aspose.Cells Workbook to PDF with 50% JPEG Compression (C#)
// Description: Creates a workbook, adds sample data, configures PdfSaveOptions to resample images at 96 dpi with JPEG quality set to 50 percent, and saves the result as a PDF file.
// Keywords: Aspose.Cells | PdfSaveOptions | JPEG compression | image resample | 50% quality | C# | .NET | Excel to PDF | reduce PDF size | SetImageResample
// Common Searches: Aspose.Cells set JPEG quality when saving PDF | PdfSaveOptions image resample C# example | Convert Excel workbook to PDF with compressed images | How to lower PDF file size using Aspose.Cells | C# Aspose.Cells export to PDF with low image quality
// Developer Intent: Generate a PDF from an Excel workbook while compressing all embedded images to JPEG at a 50 percent quality level.
// Use Cases: Produce lightweight PDF reports from Excel data for faster web delivery. | Batch‑convert multiple workbooks to PDFs with a consistent image compression setting to meet storage limits. | Create PDFs for mobile apps where bandwidth is limited, ensuring images are down‑sampled and compressed.
// AI Prompts: Show how to change the JPEG quality to 75 % and the DPI to 150 using PdfSaveOptions in Aspose.Cells. | Write a reusable C# method that accepts input workbook and output PDF paths and applies 50 % JPEG compression. | Explain the impact of SetImageResample on image scaling, resolution, and file size in the generated PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook, adds sample data, configures PdfSaveOptions to resample images at 96 dpi with JPEG quality set to 50 percent, and saves the result as a PDF file.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Text");
        sheet.Cells["B2"].PutValue(123.45);
        sheet.Cells["C3"].PutValue(DateTime.Now);

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set desired PPI (e.g., 96) and JPEG quality to 50%
        // This will convert all images in the PDF to JPEG with 50% quality
        pdfOptions.SetImageResample(96, 50);

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
