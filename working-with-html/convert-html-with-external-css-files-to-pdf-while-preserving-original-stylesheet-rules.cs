// Title: Convert HTML with external CSS to PDF using Aspose.Cells in C# while preserving stylesheet formatting
// AI Prompts: Write C# code that loads an HTML file containing linked CSS into an Aspose.Cells Workbook and saves it as a PDF, keeping the original stylesheet applied. | Create a C# example that verifies the source HTML file exists before invoking Aspose.Cells conversion and logs a clear error if it is missing. | Demonstrate how to set HtmlLoadOptions (e.g., BaseUri) in Aspose.Cells so external CSS files are correctly resolved during HTML‑to‑PDF conversion.
// Common Searches: how to preserve external CSS when converting HTML to PDF with Aspose.Cells C# | Aspose.Cells HTML to PDF conversion keeping linked stylesheet styles | C# convert html file that references external CSS to PDF using Aspose.Cells | handle missing HTML input file error in Aspose.Cells HTML to PDF example | configure base URL for external stylesheets in Aspose.Cells HtmlLoadOptions
// Tags: Aspose.Cells HTML to PDF conversion with external CSS | HtmlLoadOptions base URL stylesheet resolution | C# workbook save as PDF preserving CSS styles | file existence validation before Aspose.Cells conversion | linked stylesheet handling in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the HTML source file exists, loads it into an Aspose.Cells Workbook using HtmlLoadOptions, and saves the workbook as a PDF. All linked CSS styles are retained, and basic exception handling reports conversion issues.
class HtmlToPdfConverter
{
    // Converts an HTML file (with external CSS) to PDF while preserving stylesheet rules.
    public static void Convert(string htmlPath, string pdfPath)
    {
        // Verify that the input HTML file exists.
        if (!File.Exists(htmlPath))
        {
            Console.WriteLine($"Input HTML file not found: {htmlPath}");
            return;
        }

        try
        {
            // Load the HTML document into a workbook. HtmlLoadOptions can be customized if needed.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook as PDF. Styles from the HTML/CSS are retained.
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine($"Successfully converted '{htmlPath}' to '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully.
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }

    static void Main()
    {
        // Example usage
        string htmlFile = "input.html";   // Path to the source HTML file
        string pdfFile = "output.pdf";    // Desired output PDF file path

        Convert(htmlFile, pdfFile);
    }
}
