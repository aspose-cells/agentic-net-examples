// Title: Aspose.Cells C# – Configure PdfSaveOptions.ImageResample to 150 DPI (80% JPEG) for smaller PDFs
// Description: Load an Excel workbook, create a PdfSaveOptions object, and call SetImageResample(150, 80) to down‑sample images to screen‑quality DPI with 80 % JPEG compression. The workbook is then saved as a PDF, delivering a reduced file size while preserving visual clarity of embedded graphics.
// Keywords: Aspose.Cells PdfSaveOptions ImageResample | 150 DPI PDF export | C# reduce PDF file size | Excel to PDF image compression | 80% JPEG quality Aspose.Cells | Aspose.Cells PDF optimization | .NET PDF image resampling
// Common Searches: Aspose.Cells set image resample DPI | PdfSaveOptions 150 DPI example | how to shrink PDF size with Aspose.Cells | C# export Excel to PDF with image compression | adjust JPEG quality when saving PDF in Aspose.Cells
// Developer Intent: Apply image resampling at 150 DPI with 80 % JPEG quality during Excel‑to‑PDF conversion to lower the output size without noticeable loss of image detail.
// Use Cases: Generating compact PDF reports from image‑heavy spreadsheets for web portals. | Creating email‑friendly PDFs where attachment size limits are strict. | Automating batch conversion of large workbooks to PDFs for archival storage with controlled file size.
// AI Prompts: Show C# code that uses Aspose.Cells PdfSaveOptions.SetImageResample to export an Excel file to PDF at 150 DPI and 80 % JPEG quality. | Explain how DPI and JPEG quality settings affect PDF size and image clarity in Aspose.Cells. | Give step‑by‑step instructions to reduce PDF size by down‑sampling images when converting Excel to PDF with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load an Excel workbook, create a PdfSaveOptions object, and call SetImageResample(150, 80) to down‑sample images to screen‑quality DPI with 80 % JPEG compression. The workbook is then saved as a PDF, delivering a reduced file size while preserving visual clarity of embedded graphics.
class Program
{
    static void Main()
    {
        // Load an existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options instance
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure image resampling to 150 DPI (screen quality) with 80% JPEG quality
        // This reduces PDF size while keeping acceptable image clarity
        pdfSaveOptions.SetImageResample(150, 80);

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfSaveOptions);
    }
}
