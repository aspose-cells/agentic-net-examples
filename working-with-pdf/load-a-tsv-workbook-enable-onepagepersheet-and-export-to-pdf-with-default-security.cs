// Title: C# – Convert TSV to PDF with OnePagePerSheet using Aspose.Cells
// Description: Loads a tab‑separated values (TSV) file into an Aspose.Cells Workbook, sets PdfSaveOptions.OnePagePerSheet to true so each worksheet fits on a single PDF page, and saves the result as a PDF with the default (no password) security.
// Keywords: Aspose.Cells | TSV to PDF | OnePagePerSheet | PdfSaveOptions | C# | .NET | tab‑separated values | export workbook as PDF | Aspose.Cells PDF export
// Common Searches: Aspose.Cells convert TSV to PDF C# | OnePagePerSheet option in PdfSaveOptions | C# load TSV file with Aspose.Cells | Export workbook to PDF without password Aspose.Cells | How to render each sheet on one PDF page using Aspose.Cells
// Developer Intent: Create a PDF from a TSV workbook where every worksheet is rendered on a single page, using the default (no‑password) security settings.
// Use Cases: Generate printable reports from TSV data sources with a consistent one‑page‑per‑sheet layout. | Batch‑process multiple TSV files into PDFs for archival or distribution. | Provide a web service that accepts TSV uploads and returns a ready‑to‑print PDF without requiring password protection.
// AI Prompts: Show how to add password protection while keeping OnePagePerSheet enabled. | Give an example of customizing page margins and orientation for TSV‑to‑PDF conversion. | Suggest memory‑efficient techniques for converting very large TSV files to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTsvToPdf
{
    // Loads a tab‑separated values (TSV) file into an Aspose.Cells Workbook, sets PdfSaveOptions.OnePagePerSheet to true so each worksheet fits on a single PDF page, and saves the result as a PDF with the default (no password) security.
    class Program
    {
        static void Main()
        {
            // Load the TSV file into a workbook
            // Aspose.Cells automatically detects the TSV format based on the file extension
            Workbook workbook = new Workbook("input.tsv");

            // Create PDF save options and enable OnePagePerSheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.OnePagePerSheet = true; // All content of each sheet will be placed on a single PDF page

            // Save the workbook as PDF using the configured options
            // No security options are set, so default (no password) security is applied
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
