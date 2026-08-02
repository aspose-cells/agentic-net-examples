// Title: C# – Convert Excel Workbook to PDF with 50% JPEG Compression using Aspose.Cells
// Description: Shows how to create or load an Excel workbook, configure PdfSaveOptions to resample images at 96 dpi with JPEG quality set to 50 percent, and save the result as output_quality_50.pdf via Aspose.Cells for .NET.
// Keywords: Aspose.Cells | PdfSaveOptions | SetImageResample | JPEG compression | image quality 50 | Excel to PDF .NET | C# PDF conversion | reduce PDF size | resample images | Aspose.Cells example
// Common Searches: Aspose.Cells set JPEG quality when saving PDF | C# PdfSaveOptions image resample example | How to reduce PDF size from Excel using Aspose.Cells | Convert Excel to PDF with image compression .NET | Set image DPI and quality in Aspose.Cells PDF export
// Developer Intent: Export an Excel workbook to PDF while applying JPEG compression at a 50 % quality level.
// Use Cases: Generating lightweight PDF reports from Excel data for email distribution | Providing downloadable PDFs from web applications with controlled image size | Batch converting multiple spreadsheets to PDFs with consistent compression settings | Archiving Excel files as PDFs with reduced storage footprint
// AI Prompts: Provide C# code that loads an existing .xlsx file and saves it as a PDF using Aspose.Cells with JPEG quality set to 50 % and 96 dpi. | Explain the effect of PdfSaveOptions.SetImageResample on image resolution and compression in PDF output. | Show how to adjust both DPI and quality parameters for image resampling in Aspose.Cells PDF export. | Give a step‑by‑step guide to batch process a folder of Excel files into compressed PDFs using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for PdfSaveOptions

namespace AsposeCellsPdfConversion
{
    // Shows how to create or load an Excel workbook, configure PdfSaveOptions to resample images at 96 dpi with JPEG quality set to 50 percent, and save the result as output_quality_50.pdf via Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B2"].PutValue(123.45);
            sheet.Cells["C3"].PutValue(DateTime.Now);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set JPEG quality to 50% (and a reasonable PPI, e.g., 96)
            // This will resample images to JPEG with the specified quality.
            pdfOptions.SetImageResample(96, 50);

            // Save the workbook as a PDF using the configured options
            workbook.Save("output_quality_50.pdf", pdfOptions);

            Console.WriteLine("Workbook successfully saved as PDF with JPEG quality set to 50%.");
        }
    }
}
