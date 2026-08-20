// Title: C# – Convert HTML to PDF/A‑1b with Aspose.Cells (ICC profile not supported)
// Description: This example shows how to verify an HTML file, load it into an Aspose.Cells Workbook, configure PdfSaveOptions for PDF/A‑1b compliance, and save the result as a PDF. It includes robust error handling and notes that Aspose.Cells currently cannot embed a custom ICC color profile.
// Keywords: Aspose.Cells | C# HTML to PDF | PDF/A-1b | ICC profile | color management | PdfSaveOptions | convert HTML workbook | embed ICC | PDF conversion | error handling
// Common Searches: convert html to pdf/a-1b using Aspose.Cells C# | Aspose.Cells embed custom ICC profile in PDF | C# load html file into Aspose.Cells workbook | how to handle missing html file when converting to PDF with Aspose | Aspose.Cells PDF/A compliance options
// Developer Intent: Create a PDF/A‑1b document from an HTML source with Aspose.Cells for .NET while acknowledging the current lack of ICC profile embedding support.
// Use Cases: Generate archival‑ready PDFs from web‑based reports that must meet PDF/A‑1b standards. | Validate the presence of the source HTML file before conversion to avoid runtime errors. | Prepare code for future ICC profile support by adding a clear placeholder comment.
// AI Prompts: Write C# code that converts an HTML file to PDF/A‑2b with Aspose.Cells and embeds a custom ICC profile, assuming the API exists. | Explain how to monitor Aspose.Cells release notes for upcoming ICC profile embedding features and update the sample accordingly. | Create unit tests for the HTML‑to‑PDF conversion routine that cover missing file detection, successful conversion, and unexpected exception handling.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example shows how to verify an HTML file, load it into an Aspose.Cells Workbook, configure PdfSaveOptions for PDF/A‑1b compliance, and save the result as a PDF. It includes robust error handling and notes that Aspose.Cells currently cannot embed a custom ICC color profile.
class HtmlToPdfWithIcc
{
    static void Main()
    {
        try
        {
            const string htmlPath = "sample.html";
            const string pdfPath = "output.pdf";

            // Verify that the source HTML file exists to avoid FileNotFoundException.
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                return;
            }

            // Load the HTML file into a workbook using the appropriate constructor.
            Workbook workbook = new Workbook(htmlPath);

            // Create PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set PDF/A compliance for better color management.
                Compliance = PdfCompliance.PdfA1b
            };

            // NOTE: Aspose.Cells does not currently expose a direct API to embed an ICC profile.
            // If future versions provide such functionality, it can be set here.

            // Save the workbook as a PDF file with the specified options.
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"PDF successfully created at '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
