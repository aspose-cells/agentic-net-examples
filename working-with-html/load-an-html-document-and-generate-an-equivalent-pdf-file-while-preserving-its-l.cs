using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace HtmlToPdfConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source HTML file
            string htmlFilePath = "input.html";

            // Load the HTML document into a Workbook object
            // Aspose.Cells can directly parse HTML and create an Excel workbook representation
            Workbook workbook = new Workbook(htmlFilePath);

            // Configure PDF save options to preserve layout and styling
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure that the document structure (e.g., headings) is retained
                ExportDocumentStructure = true,

                // Use PDF/A-1b compliance for better compatibility (optional)
                Compliance = PdfCompliance.PdfA1b,

                // Calculate any formulas before saving (if present)
                CalculateFormula = true
            };

            // Path for the output PDF file
            string pdfFilePath = "output.pdf";

            // Save the workbook as a PDF using the configured options
            workbook.Save(pdfFilePath, pdfOptions);

            Console.WriteLine($"HTML file '{htmlFilePath}' has been successfully converted to PDF '{pdfFilePath}'.");
        }
    }
}