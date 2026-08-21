// Title: Convert HTML to PDF with 300 DPI using Aspose.Cells for .NET
// Description: Learn how to load an HTML file into an Aspose.Cells Workbook, set the global DPI to 300, resample images at 300 PPI with maximum JPEG quality, apply standard PDF optimization, and save a high‑resolution PDF. Ideal for print‑ready documents and sharp graphics.
// Keywords: Aspose.Cells HTML to PDF | 300 DPI PDF export .NET | PdfSaveOptions image resample | high quality PDF Aspose.Cells | set global DPI Aspose.Cells | standard PDF optimization | C# convert HTML to PDF | Aspose.Cells rendering DPI
// Common Searches: Aspose.Cells convert HTML to PDF with 300 DPI | set DPI for PDF export in Aspose.Cells C# | how to resample images when saving PDF with Aspose.Cells | high‑resolution PDF from HTML using Aspose.Cells | C# code sample for HTML to PDF 300 DPI
// Developer Intent: Create a PDF from an HTML workbook with 300 DPI graphics to ensure print‑quality output.
// Use Cases: Generate print‑ready PDFs from web‑based reports while preserving image sharpness. | Batch‑process HTML invoices into high‑resolution PDFs for archival compliance. | Export HTML dashboards to marketing PDFs where image clarity is critical.
// AI Prompts: Provide a C# example that embeds custom fonts while keeping the 300 DPI setting for HTML‑to‑PDF conversion with Aspose.Cells. | Show how to write the high‑resolution PDF to a MemoryStream instead of a file, preserving all DPI and optimization options.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Learn how to load an HTML file into an Aspose.Cells Workbook, set the global DPI to 300, resample images at 300 PPI with maximum JPEG quality, apply standard PDF optimization, and save a high‑resolution PDF. Ideal for print‑ready documents and sharp graphics.
class HtmlToPdfHighDpi
{
    static void Main()
    {
        // Set the global DPI to 300 for high‑quality graphics
        CellsHelper.DPI = 300;

        // Load the HTML file into a workbook
        Workbook workbook = new Workbook("input.html");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Resample all images to 300 PPI and keep maximum JPEG quality
        pdfOptions.SetImageResample(300, 100);

        // Use standard optimization for high print quality
        pdfOptions.OptimizationType = PdfOptimizationType.Standard;

        // Save the workbook as a PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
