// Title: Aspose.Cells C# – Convert Excel to PDF with 80% JPEG Image Quality
// Description: Demonstrates how to load or create a workbook, configure PdfSaveOptions with 150 DPI and 80 % JPEG quality via SetImageResample, and save the file as a PDF that balances file size and visual clarity.
// Keywords: Aspose.Cells | C# PDF conversion | Excel to PDF | SetImageResample | image quality 80% | JPEG compression | 150 DPI | PdfSaveOptions | workbook export .NET | balanced file size
// Common Searches: Aspose.Cells set JPEG quality when saving PDF | C# PdfSaveOptions SetImageResample example | export Excel as PDF with reduced image size | adjust DPI and compression in Aspose.Cells PDF output | how to limit PDF file size from Excel in .NET
// Developer Intent: Generate a PDF from an Excel workbook while applying 80 % JPEG compression to keep the document lightweight without sacrificing readability.
// Use Cases: Create email‑ready PDF reports from spreadsheets with controlled image size. | Batch‑process multiple workbooks to PDF using a uniform 80 % image quality setting for storage limits. | Export charts and embedded pictures to PDF while preserving clarity through DPI and compression tuning.
// AI Prompts: Show a C# snippet that loads an Excel file, sets PdfSaveOptions to 150 DPI and 80 % JPEG quality, and saves it as PDF using Aspose.Cells. | Explain how SetImageResample affects PDF output size and image clarity in Aspose.Cells. | Provide step‑by‑step instructions to convert a workbook to PDF with custom image quality parameters in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to load or create a workbook, configure PdfSaveOptions with 150 DPI and 80 % JPEG quality via SetImageResample, and save the file as a PDF that balances file size and visual clarity.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for PDF conversion");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set image resampling: desired PPI (e.g., 150) and JPEG quality to 80%
        pdfOptions.SetImageResample(150, 80);

        // Save the workbook as PDF with the specified image quality
        workbook.Save("output.pdf", pdfOptions);
    }
}
