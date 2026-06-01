using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class HtmlToPdfWithAccessibility
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlFilePath = "input.html";

            // Verify that the HTML file exists
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlFilePath}'.");
                return;
            }

            // Load the HTML content into a workbook
            Workbook workbook = new Workbook(htmlFilePath);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Configure security options to enable accessibility extraction
            PdfSecurityOptions security = new PdfSecurityOptions
            {
                // Allow screen readers to extract text and graphics
                AccessibilityExtractContent = true
            };
            pdfOptions.SecurityOptions = security;

            // Optional: set PDF compliance level (e.g., PDF/A-1b for better accessibility)
            // Uncomment the following line if the PdfCompliance enum is available in your Aspose.Cells version
            // pdfOptions.Compliance = PdfCompliance.PdfA1b;

            // Save the workbook as a PDF with the specified options
            string pdfOutputPath = "output.pdf";
            workbook.Save(pdfOutputPath, pdfOptions);

            Console.WriteLine($"HTML file '{htmlFilePath}' has been converted to PDF with accessibility tags at '{pdfOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}