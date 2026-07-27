// Title: Convert HTML with External CSS to PDF in C# using Aspose.Cells
// Description: Shows how to load an HTML file that references external CSS files into an Aspose.Cells Workbook via HtmlLoadOptions and export it to PDF, keeping the original stylesheet formatting intact.
// Keywords: Aspose.Cells | HTML to PDF conversion | C# | .NET | external CSS | HtmlLoadOptions | SaveFormat.Pdf | preserve CSS styling | workbook export | styled HTML PDF
// Common Searches: Aspose.Cells convert HTML with linked CSS to PDF | C# preserve CSS when converting HTML to PDF | HtmlLoadOptions external stylesheet PDF output | How to keep HTML styles in PDF using Aspose.Cells | Save HTML page as PDF with original CSS in .NET
// Developer Intent: Generate a PDF from an HTML document that uses external CSS files while retaining the exact visual appearance.
// Use Cases: Automate PDF invoice generation from HTML templates that include linked style sheets. | Batch‑process marketing web pages into PDFs for archival without losing design fidelity. | Embed styled HTML email previews as PDF attachments in a .NET application.
// AI Prompts: Write C# code with Aspose.Cells to convert an HTML file that links external CSS into a PDF, ensuring all styles are applied. | Explain how HtmlLoadOptions and Workbook.Save with SaveFormat.Pdf preserve external CSS during HTML‑to‑PDF conversion. | Suggest robust error handling for converting multiple HTML files with linked stylesheets to PDFs using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to load an HTML file that references external CSS files into an Aspose.Cells Workbook via HtmlLoadOptions and export it to PDF, keeping the original stylesheet formatting intact.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the HTML file that references external CSS files
        string htmlFilePath = "input.html";

        // Path where the resulting PDF will be saved
        string pdfFilePath = "output.pdf";

        try
        {
            // Ensure the HTML input file exists
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"Input HTML file not found: {htmlFilePath}");
                return;
            }

            // Load the HTML document into a workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Save the workbook as PDF. The visual appearance, including CSS styling,
            // is preserved in the generated PDF.
            workbook.Save(pdfFilePath, SaveFormat.Pdf);

            Console.WriteLine("HTML has been successfully converted to PDF with original stylesheet rules.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
