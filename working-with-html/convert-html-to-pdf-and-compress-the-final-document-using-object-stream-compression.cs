// Title: C# – Convert HTML to PDF with Flate object‑stream compression using Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions for Flate compression and MinimumSize optimization, and saves a compact PDF suitable for email or archival storage.
// Keywords: Aspose.Cells | HTML to PDF | C# | Flate compression | PdfSaveOptions | object stream compression | minimum size PDF | PDF optimization | .NET PDF compression | Aspose.Cells PDF
// Common Searches: Aspose.Cells convert HTML to PDF C# | Flate object stream compression PDF Aspose | How to reduce PDF size with Aspose.Cells | PdfSaveOptions compression options .NET | Compress PDF generated from HTML using Aspose
// Developer Intent: The developer wants to turn an HTML document into a PDF and apply Flate object‑stream compression to minimize the resulting file size.
// Use Cases: Create email‑friendly PDFs from HTML templates. | Archive web pages as small PDFs for long‑term storage. | Batch‑process HTML reports into compressed PDFs for distribution.
// AI Prompts: Generate C# code that converts an HTML file to a PDF with Flate compression using Aspose.Cells. | Explain the impact of PdfCompressionCore.Flate and PdfOptimizationType.MinimumSize on PDF size in Aspose.Cells. | Show how to enable additional PDF options (e.g., font embedding) while keeping object‑stream compression active.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions for Flate compression and MinimumSize optimization, and saves a compact PDF suitable for email or archival storage.
class HtmlToPdfCompressed
{
    static void Main()
    {
        // Load the HTML file into a workbook
        Workbook workbook = new Workbook("input.html");

        // Configure PDF save options with object stream compression
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.PdfCompression = PdfCompressionCore.Flate; // Apply Flate compression to PDF streams
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize; // Optional: minimize file size

        // Save the workbook as a compressed PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
