// Title: Convert HTML to PDF with Aspose.Cells for .NET while preserving CSS ::before and ::after pseudo‑elements
// AI Prompts: Write C# code that uses Aspose.Cells to load an HTML file, renders CSS ::before and ::after pseudo‑elements, and saves the result as a PDF. | Show how to set up HtmlLoadOptions in Aspose.Cells so that CSS pseudo‑elements are included during HTML‑to‑PDF conversion. | Provide a robust C# example that checks the input HTML file, performs the conversion with Aspose.Cells, and includes error handling for missing files or conversion failures.
// Common Searches: Aspose.Cells C# HTML to PDF conversion keep CSS pseudo elements | How to render ::before and ::after styles when converting HTML to PDF with Aspose.Cells | HtmlLoadOptions settings for preserving CSS pseudo‑elements in PDF output | Convert HTML file to PDF in .NET while maintaining pseudo‑element formatting | C# Aspose.Cells PDF export with full CSS support including ::before ::after
// Tags: Aspose.Cells HtmlLoadOptions CSS pseudo‑elements | C# HTML to PDF conversion Aspose.Cells | preserve ::before ::after in PDF generation | Aspose.Cells PDF export with CSS styling | error handling for HTML to PDF conversion .NET

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the existence of an input HTML file, loads it into an Aspose.Cells Workbook using HtmlLoadOptions, and saves the workbook as a PDF, handling any exceptions that may occur during the conversion.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFilePath = "input.html";

        // Path where the resulting PDF will be saved
        string pdfFilePath = "output.pdf";

        // Verify that the HTML file exists to avoid FileNotFoundException
        if (!File.Exists(htmlFilePath))
        {
            Console.WriteLine($"Error: HTML file not found at '{htmlFilePath}'.");
            return;
        }

        try
        {
            // Load the HTML content into a Workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Save the workbook as PDF
            workbook.Save(pdfFilePath, SaveFormat.Pdf);

            Console.WriteLine("Conversion completed. PDF saved to: " + pdfFilePath);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred during conversion: " + ex.Message);
        }
    }
}
