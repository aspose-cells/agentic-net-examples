// Title: C# – Convert HTML to PDF/A‑1b with Aspose.Cells
// Description: This example shows how to load an HTML file into an Aspose.Cells Workbook, configure PdfSaveOptions for PDF/A‑1b compliance, and save the result as an archival‑grade PDF using .NET.
// Keywords: Aspose.Cells HTML to PDF conversion | PDF/A-1b compliance .NET | C# convert HTML to archival PDF | PdfSaveOptions Compliance property | Generate PDF/A-1b from HTML workbook
// Common Searches: Convert HTML to PDF/A-1b using Aspose.Cells C# | Aspose.Cells set PDF compliance to PDF/A-1b | C# archive HTML reports as PDF/A-1b | Save HTML workbook as PDF/A-1b with Aspose.Cells
// Developer Intent: Create a PDF/A‑1b document from an HTML source with Aspose.Cells in C#.
// Use Cases: Preserve web‑based financial statements for regulatory filing. | Store HTML invoices as long‑term, standards‑compliant PDFs. | Automate batch conversion of HTML assets to PDF/A‑1b in a background service.
// AI Prompts: Write C# code that converts a folder of HTML files to PDF/A‑1b using Aspose.Cells with comprehensive error handling. | Explain how to embed custom fonts and add document metadata when saving PDF/A‑1b with Aspose.Cells. | Show how to tweak PdfSaveOptions (e.g., image quality, compression) for optimal PDF/A‑1b output in a console app.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example shows how to load an HTML file into an Aspose.Cells Workbook, configure PdfSaveOptions for PDF/A‑1b compliance, and save the result as an archival‑grade PDF using .NET.
class HtmlToPdfA1bConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFile = "input.html";

        // Path where the PDF/A‑1b file will be saved
        string pdfFile = "output.pdf";

        // Load the HTML file into a workbook
        // The Workbook constructor automatically detects the format based on the file extension
        Workbook workbook = new Workbook(htmlFile);

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set compliance level to PDF/A‑1b for long‑term archival
        pdfOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the workbook as a PDF with the specified compliance
        workbook.Save(pdfFile, pdfOptions);

        Console.WriteLine("HTML has been successfully converted to PDF/A‑1b.");
    }
}
