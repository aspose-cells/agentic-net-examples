// Title: Export Excel to PDF/A‑2b with original color profiles – Aspose.Cells for .NET
// Description: Demonstrates how to use Aspose.Cells for .NET to save a workbook as PDF/A‑2b. By setting PdfSaveOptions.Compliance to PdfCompliance.PdfA2b, Aspose.Cells automatically retains any embedded ICC color profiles, delivering accurate color reproduction without extra configuration.
// Keywords: Aspose.Cells PDF/A-2b | C# PDF/A-2b export | preserve ICC color profile | PdfSaveOptions Compliance | Excel to PDF/A-2b | color fidelity PDF export | Aspose.Cells .NET PDF compliance | archival PDF from Excel
// Common Searches: Aspose.Cells set PDF/A-2b compliance C# | keep ICC color profile when exporting Excel to PDF with Aspose | C# convert workbook to PDF/A-2b preserving colors | PDF/A-2b export options Aspose.Cells | how to retain original color profiles in PDF generated from Excel
// Developer Intent: Configure Aspose.Cells to generate a PDF/A‑2b file from an Excel workbook while maintaining the workbook’s embedded color profiles.
// Use Cases: Create archival‑grade PDFs from financial reports that must meet PDF/A‑2b standards and keep exact brand colors. | Develop a server‑side batch job that converts dozens of spreadsheets to PDF/A‑2b for legal compliance, preserving embedded ICC profiles. | Expose a REST API that returns PDF/A‑2b documents with original color fidelity for downstream printing workflows.
// AI Prompts: Generate C# code using Aspose.Cells to convert an existing workbook to PDF/A‑2b and retain its ICC color profile. | Explain whether additional settings are required in Aspose.Cells to preserve color profiles when saving as PDF/A‑2b. | Provide a reusable method that loads a workbook, applies PdfSaveOptions with PdfCompliance.PdfA2b, and saves the PDF while ensuring color accuracy.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to use Aspose.Cells for .NET to save a workbook as PDF/A‑2b. By setting PdfSaveOptions.Compliance to PdfCompliance.PdfA2b, Aspose.Cells automatically retains any embedded ICC color profiles, delivering accurate color reproduction without extra configuration.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("PDF/A‑2b compliance with original color profiles");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set the PDF/A‑2b compliance level
            Compliance = PdfCompliance.PdfA2b
            // Aspose.Cells preserves the original color profiles by default when saving to PDF.
            // No additional settings are required for color profile preservation.
        };

        // Save the workbook as a PDF with the specified compliance
        workbook.Save("Output_PdfA2b.pdf", pdfOptions);
    }
}
