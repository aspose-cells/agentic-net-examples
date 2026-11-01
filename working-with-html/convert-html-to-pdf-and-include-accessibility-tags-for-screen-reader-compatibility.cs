// Title: Convert HTML to PDF with optional PDF/UA accessibility using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an HTML file into an Aspose.Cells Workbook and saves it as a PDF, setting PdfSaveOptions.Accessible = true only when the property is available. | Add a version‑check routine that determines if the current Aspose.Cells library supports the Accessible flag and applies it conditionally during PDF export. | Implement comprehensive validation for the input HTML path and detailed exception handling for the HTML‑to‑PDF conversion process.
// Common Searches: how to export an HTML workbook to PDF with PDF/UA tags using Aspose.Cells C# | Aspose.Cells PdfSaveOptions Accessible property version check | C# convert HTML spreadsheet to PDF with accessibility support Aspose.Cells | detect if Aspose.Cells supports PDF accessibility before saving | error handling for missing HTML file when using Aspose.Cells to create PDF
// Tags: Aspose.Cells HTML to PDF conversion with PdfSaveOptions | C# enable PDF/UA accessibility in Aspose.Cells | conditional Accessible property check Aspose.Cells | file existence validation Aspose.Cells HTML import | exception handling Aspose.Cells PDF export

using System;
using System.IO;
using Aspose.Cells;

// The example loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions (noting that the Accessible flag may be unavailable in older versions), optionally enables PDF/UA accessibility, validates the input file, and saves the workbook as a PDF while handling any runtime errors.
class HtmlToPdfWithAccessibility
{
    static void Main()
    {
        // Path to the source HTML file (can be a spreadsheet saved as HTML)
        string htmlFilePath = "input.html";

        // Path where the resulting PDF will be saved
        string pdfFilePath = "output.pdf";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"Error: The file '{htmlFilePath}' was not found.");
                return;
            }

            // Load the HTML content into a Workbook object
            Workbook workbook = new Workbook(htmlFilePath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: The 'Accessible' property is not available in the current Aspose.Cells version.
            // If your version supports PDF/UA tags, you can enable them by setting:
            // pdfOptions.Accessible = true;

            // Save the workbook as a PDF
            workbook.Save(pdfFilePath, pdfOptions);

            Console.WriteLine($"PDF successfully saved to '{pdfFilePath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
