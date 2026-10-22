// Title: Convert an HTML file containing CSS gradient backgrounds to PDF with Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an HTML document with CSS gradient backgrounds into an Aspose.Cells Workbook using HtmlLoadOptions and saves it as a PDF. | Show how to verify the input HTML file exists and gracefully handle exceptions during the HTML‑to‑PDF conversion with Aspose.Cells. | Provide an example of customizing PDF export settings such as page size or orientation while keeping CSS gradient fills intact.
// Common Searches: how to preserve CSS gradient backgrounds when converting HTML to PDF using Aspose.Cells C# | Aspose.Cells HtmlLoadOptions gradient fill support .NET example | C# convert HTML file with background-image gradients to PDF with Aspose.Cells | error handling for missing HTML file during Aspose.Cells PDF export | change PDF page orientation while exporting HTML workbook in Aspose.Cells
// Tags: Aspose.Cells HTML to PDF gradient support | HtmlLoadOptions preserve CSS gradients | C# export workbook as PDF with background images | PDF page orientation Aspose.Cells | exception handling Aspose.Cells HTML conversion

using System;
using System.IO;
using Aspose.Cells;

// C# program that loads an HTML file (including CSS gradient backgrounds) into an Aspose.Cells Workbook via HtmlLoadOptions and saves it as a PDF, preserving the gradient fills and allowing optional PDF export customization.
class Program
{
    static void Main()
    {
        try
        {
            // Input HTML file containing CSS gradient backgrounds
            string htmlFile = "input.html";

            // Output PDF file path
            string pdfFile = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(htmlFile))
            {
                Console.WriteLine($"Input file not found: {htmlFile}");
                return;
            }

            // Load the HTML into a Workbook. HtmlLoadOptions handles CSS styles,
            // including gradient fills in background images.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Create the workbook from the HTML source
            Workbook workbook = new Workbook(htmlFile, loadOptions);

            // Save the workbook as a PDF. Aspose.Cells renders the gradient fills
            // into the PDF output, preserving the visual appearance.
            workbook.Save(pdfFile, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully saved to: {pdfFile}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
