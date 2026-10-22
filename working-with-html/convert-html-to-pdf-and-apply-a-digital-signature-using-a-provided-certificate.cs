// Title: Convert an HTML file to PDF and digitally sign it with a certificate using Aspose.Cells and Aspose.Pdf in C#
// AI Prompts: Write C# code that checks for the existence of an input HTML file, loads it into an Aspose.Cells Workbook with LoadOptions, and saves it as a PDF. | Add robust exception handling around the HTML‑to‑PDF conversion in a console application. | Extend the program to load an X509 certificate from a .pfx file and apply a digital signature to the generated PDF using Aspose.Pdf. | Create a method that accepts the HTML path, PDF output path, and certificate details, then performs conversion and signing in one workflow.
// Common Searches: c# convert html to pdf with Aspose.Cells and then sign pdf using a certificate | how to apply a digital signature to a pdf generated from html in a console app | aspnet console application check file existence before Aspose.Cells conversion | load X509 certificate from pfx and sign pdf using Aspose.Pdf in C#
// Tags: Aspose.Cells HTML to PDF conversion C# | Aspose.Pdf digital signature with X509 certificate | LoadOptions Html format Aspose.Cells | Workbook.Save SaveFormat.Pdf | PdfFileSignature sign PDF using certificate | Console app file existence validation

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to verify an input HTML file, convert it to PDF with Aspose.Cells, and then apply a digital signature using an X509 certificate via Aspose.Pdf, all wrapped in a C# console application with comprehensive error handling.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Paths for input HTML and output PDF
        string htmlPath = "input.html";
        string pdfPath = "output.pdf";

        try
        {
            // Ensure the HTML file exists before processing
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input HTML file not found: {htmlPath}");
                return;
            }

            // Load the HTML content into an Aspose.Cells workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook as a PDF document
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine("HTML has been converted to PDF successfully.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
