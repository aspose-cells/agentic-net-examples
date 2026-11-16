// Title: Convert HTML containing CSS gradient backgrounds to PDF while preserving gradients using Aspose.Cells in C#
// Common Searches: C# convert HTML with CSS gradient backgrounds to PDF using Aspose.Cells | preserve CSS gradient fills when exporting HTML to PDF with Aspose.Cells | how to use HtmlLoadOptions to retain background gradients in PDF output | Aspose.Cells HTML to PDF conversion keeping CSS styling in .NET | save HTML file as PDF with gradient backgrounds in C#
// Tags: aspocells htmlloadoptions gradient rendering | c# html to pdf conversion aspocells | preserve css gradients pdf aspocells | save workbook as pdf css backgrounds | aspocells html to pdf css styling

using System;
using System.IO;
using Aspose.Cells;

// This example loads an HTML file that includes CSS gradient backgrounds into an Aspose.Cells Workbook using HtmlLoadOptions, then saves the workbook as a PDF, ensuring the gradient fills are retained in the generated PDF.
class Program
{
    static void Main()
    {
        // Path to the source HTML file that contains CSS gradient backgrounds
        string htmlFilePath = "input.html";

        // Path where the resulting PDF will be saved
        string pdfFilePath = "output.pdf";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlFilePath))
                throw new FileNotFoundException($"Input file not found: {htmlFilePath}");

            // Load options for HTML. Default settings already handle CSS and background images.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML document into a Workbook using the specified load options.
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(pdfFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the Workbook as a PDF file. The PDF rendering engine retains CSS gradients.
            workbook.Save(pdfFilePath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully created at: {pdfFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
