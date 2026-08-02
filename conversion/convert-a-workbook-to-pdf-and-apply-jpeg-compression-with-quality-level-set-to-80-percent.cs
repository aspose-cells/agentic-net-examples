// Title: C# – Convert Excel Workbook to PDF with 80% JPEG Compression using Aspose.Cells
// Description: Demonstrates how to create or load an Aspose.Cells Workbook, set PdfSaveOptions to resample images at 96 PPI with JPEG quality of 80 %, and save the result as a PDF file (output.pdf).
// Keywords: Aspose.Cells PDF conversion | C# Excel to PDF | JPEG compression 80 percent | PdfSaveOptions SetImageResample | reduce PDF file size | 96 PPI image resampling | Aspose.Cells .NET | image quality control in PDF | US compliance PDF size | European PDF optimization
// Common Searches: Aspose.Cells set JPEG quality when saving to PDF | PdfSaveOptions image resample example C# | How to compress images in Excel‑to‑PDF conversion | Convert Excel workbook to PDF with 80% image quality | Reduce PDF size using Aspose.Cells image settings
// Developer Intent: Export an Excel workbook to PDF while applying 80 % JPEG compression to embedded images.
// Use Cases: Generate lightweight PDF reports from spreadsheets for email distribution. | Provide web users with fast‑download PDFs by lowering image quality to 80 %. | Batch‑process multiple workbooks to meet corporate PDF size policies.
// AI Prompts: Write C# code that loads an existing Excel file and saves it as a PDF with 80 % JPEG compression and 96 PPI using Aspose.Cells. | Explain the impact of PdfSaveOptions.SetImageResample on PDF size and image clarity. | Create a reusable method: (string inputPath, string outputPath) → PDF with 80 % JPEG quality via Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create or load an Aspose.Cells Workbook, set PdfSaveOptions to resample images at 96 PPI with JPEG quality of 80 %, and save the result as a PDF file (output.pdf).
class ConvertWorkbookToPdfWithJpegCompression
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Text");
        sheet.Cells["B2"].PutValue(123.45);
        sheet.Cells["C3"].PutValue(DateTime.Now);

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set image resampling: 96 PPI and JPEG quality to 80%
        pdfOptions.SetImageResample(96, 80);

        // Save the workbook as PDF using the options
        workbook.Save("output.pdf", pdfOptions);
    }
}
