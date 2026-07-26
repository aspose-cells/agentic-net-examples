// Title: Convert HTML with CSS Gradient Backgrounds to PDF using Aspose.Cells for .NET
// Description: A concise C# example that loads an HTML file containing CSS gradient backgrounds into an Aspose.Cells Workbook (LoadFormat.Html) and saves it as a PDF (SaveFormat.Pdf) while preserving the gradient fills and overall visual styling.
// Keywords: Aspose.Cells | HTML to PDF conversion | CSS gradient background | preserve gradient fills | C# | LoadOptions | LoadFormat.Html | SaveFormat.Pdf | Workbook conversion | .NET PDF rendering | gradient fill preservation
// Common Searches: Aspose.Cells convert HTML with gradients to PDF | C# preserve CSS gradient when exporting HTML to PDF | Load HTML workbook Aspose.Cells gradient background | HTML to PDF with gradient fills .NET | Aspose.Cells SaveFormat.Pdf gradient support
// Developer Intent: Generate a PDF from an HTML document that uses CSS gradient backgrounds, ensuring the gradients appear unchanged in the output, by using Aspose.Cells for .NET.
// Use Cases: Create printable PDFs from marketing web pages that rely on gradient styling without losing design fidelity. | Automate archival of email or landing‑page HTML templates as PDFs while keeping background effects intact. | Batch‑process a collection of HTML reports with gradient backgrounds into high‑quality PDFs for documentation.
// AI Prompts: Write C# code that loads an HTML file with CSS gradients into an Aspose.Cells Workbook and saves it as a PDF preserving the gradients. | Explain how LoadOptions with LoadFormat.Html affects rendering of CSS gradient backgrounds in Aspose.Cells PDF output. | Suggest best‑practice error handling and logging for converting multiple HTML files to PDF with Aspose.Cells, ensuring gradient fills remain intact.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHtmlToPdf
{
    // A concise C# example that loads an HTML file containing CSS gradient backgrounds into an Aspose.Cells Workbook (LoadFormat.Html) and saves it as a PDF (SaveFormat.Pdf) while preserving the gradient fills and overall visual styling.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains CSS gradient background images
            string htmlPath = "input.html";

            // Path for the resulting PDF file
            string pdfPath = "output.pdf";

            try
            {
                // Verify that the HTML file exists to avoid FileNotFoundException
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: HTML file '{htmlPath}' not found.");
                    return;
                }

                // Load the HTML file into the workbook using LoadOptions with Html format
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(htmlPath, loadOptions);

                // Save the workbook as PDF, preserving visual elements such as gradient fills
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"HTML file '{htmlPath}' has been converted to PDF '{pdfPath}' with gradient fills preserved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
