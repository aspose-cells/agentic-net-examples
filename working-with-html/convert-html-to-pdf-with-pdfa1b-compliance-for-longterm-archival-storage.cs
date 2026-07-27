// Title: C# – Convert HTML to PDF/A‑1b with Aspose.Cells
// Description: Load an HTML file into an Aspose.Cells Workbook, configure PdfSaveOptions.Compliance to PdfA1b, and save the workbook as a PDF/A‑1b file ready for long‑term archival.
// Keywords: Aspose.Cells | HTML to PDF | PDF/A-1b | .NET | C# conversion | PdfSaveOptions | archival PDF | PDF compliance | document preservation | batch conversion
// Common Searches: Aspose.Cells HTML to PDF/A-1b C# | how to create PDF/A-1b from HTML .NET | set PDF compliance Aspose.Cells | convert web page to archival PDF using C# | batch HTML to PDF/A-1b Aspose.Cells
// Developer Intent: Create a PDF/A‑1b compliant PDF from an HTML source using Aspose.Cells in C#.
// Use Cases: Archiving web‑based reports for legal retention | Converting HTML invoices to PDF/A‑1b for regulatory compliance | Batch processing of HTML documents into archival PDFs within a .NET service | Generating PDF/A‑1b attachments for automated email workflows
// AI Prompts: Write C# code that loads an HTML file into an Aspose.Cells Workbook and saves it as a PDF/A‑1b compliant PDF. | Explain how to configure PdfSaveOptions for PDF/A‑1b compliance, including font embedding and metadata settings. | Show a loop that converts a list of HTML files to PDF/A‑1b using Aspose.Cells, handling errors and logging progress. | Suggest ways to optimize the conversion speed when processing large batches of HTML documents to PDF/A‑1b.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an HTML file into an Aspose.Cells Workbook, configure PdfSaveOptions.Compliance to PdfA1b, and save the workbook as a PDF/A‑1b file ready for long‑term archival.
class HtmlToPdfA1b
{
    static void Main()
    {
        // Load the source HTML file into a workbook.
        // Aspose.Cells can directly load HTML documents.
        Workbook workbook = new Workbook("input.html");

        // Create PDF save options and set the compliance level to PDF/A‑1b.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the workbook as a PDF file using the specified options.
        workbook.Save("output.pdf", pdfOptions);
    }
}
