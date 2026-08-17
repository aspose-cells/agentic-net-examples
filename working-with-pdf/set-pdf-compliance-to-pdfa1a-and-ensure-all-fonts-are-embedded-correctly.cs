// Title: Create PDF/A‑1a compliant PDFs with fully embedded fonts using Aspose.Cells for C#
// Description: Demonstrates how to configure Aspose.Cells PdfSaveOptions to produce a PDF/A‑1a file that embeds all standard Windows fonts, uses Identity encoding, and sets Arial as the default font before saving an Excel workbook as a PDF.
// Keywords: Aspose.Cells | C# | .NET | PDF/A-1a | embed fonts | PdfSaveOptions | PDF compliance | Identity font encoding | DefaultFont Arial | Excel to PDF export | archival PDF
// Common Searches: Aspose.Cells PDF/A-1a C# example | how to embed fonts in PDF with Aspose.Cells .NET | set PDF compliance to PDF/A-1a using PdfSaveOptions | default font Arial PDF export Aspose.Cells | generate archival PDF from Excel in C#
// Developer Intent: The developer needs to save an Excel workbook as a PDF/A‑1a document with every font embedded for reliable, standards‑compliant rendering.
// Use Cases: Archiving financial reports that must meet PDF/A‑1a standards and retain exact visual fidelity. | Distributing multilingual Excel‑derived PDFs where consistent font rendering is critical. | Creating regulatory‑compliant documents that require all fonts to be embedded and encoded with Identity.
// AI Prompts: Show how to switch the compliance level to PDF/A‑2b while keeping font embedding enabled. | Provide code that saves each worksheet of a workbook to separate PDF/A‑1a files with embedded fonts. | Explain methods to programmatically verify PDF/A‑1a compliance and confirm that all fonts are embedded.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to configure Aspose.Cells PdfSaveOptions to produce a PDF/A‑1a file that embeds all standard Windows fonts, uses Identity encoding, and sets Arial as the default font before saving an Excel workbook as a PDF.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample content
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample text with PDF/A‑1a compliance and embedded fonts.");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set compliance level to PDF/A‑1a
            Compliance = PdfCompliance.PdfA1a,
            // Ensure all standard Windows fonts are embedded
            EmbedStandardWindowsFonts = true,
            // Use Identity encoding for all embedded fonts
            FontEncoding = PdfFontEncoding.Identity,
            // Specify a default font to handle Unicode characters
            DefaultFont = "Arial"
        };

        // Save the workbook as a PDF with the specified options
        workbook.Save("Output_PdfA1a.pdf", pdfOptions);
    }
}
