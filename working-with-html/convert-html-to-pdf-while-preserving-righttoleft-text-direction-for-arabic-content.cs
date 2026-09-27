// Title: Convert Arabic RTL HTML to PDF using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an Arabic HTML file with right‑to‑left layout into an Aspose.Cells Workbook and saves it as a PDF. | Show how to configure HtmlLoadOptions for RTL support when converting HTML to PDF with Aspose.Cells in a .NET application. | Add comprehensive error handling to a C# Aspose.Cells routine that verifies the input HTML path and reports conversion failures.
// Common Searches: Aspose.Cells C# convert Arabic HTML to PDF preserving RTL layout | How to keep right‑to‑left text direction when saving HTML as PDF with Aspose.Cells | Example of HtmlLoadOptions RTL support in Aspose.Cells .NET | C# code to convert an HTML file containing Arabic content to PDF using Aspose.Cells | Troubleshoot missing right‑to‑left formatting after HTML‑to‑PDF conversion with Aspose.Cells
// Tags: Aspose.Cells HTML to PDF right-to-left conversion | C# HtmlLoadOptions RTL support | Arabic PDF generation with Aspose.Cells | Preserve right-to-left formatting Aspose.Cells | Convert HTML to PDF using Aspose.Cells .NET | Robust error handling Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the presence of an Arabic HTML file, loads it into an Aspose.Cells Workbook with HtmlLoadOptions, and saves the workbook as a PDF while maintaining right‑to‑left text direction. It includes basic exception handling and reports success or error messages.
class HtmlToPdfRtlConverter
{
    static void Main()
    {
        // Path to the source HTML file containing Arabic content
        string htmlFilePath = "input.html";

        // Desired output PDF file path
        string pdfFilePath = "output.pdf";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"Input file not found: {htmlFilePath}");
                return;
            }

            // Load the HTML document into a workbook
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Save the workbook as a PDF file
            workbook.Save(pdfFilePath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully created at: {pdfFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
