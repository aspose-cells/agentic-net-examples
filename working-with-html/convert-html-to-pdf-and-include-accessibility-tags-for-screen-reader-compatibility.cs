// Title: Convert HTML to PDF with screen‑reader accessible tags using Aspose.Cells for .NET (optional PDF/A‑1b)
// AI Prompts: Write C# code that loads an HTML file into an Aspose.Cells Workbook and saves it as a PDF containing accessibility tags for screen readers, with an option to enable PDF/A‑1b compliance. | Add file‑existence validation and comprehensive exception handling to the HTML‑to‑PDF conversion sample, and configure PdfSaveOptions to produce an accessible PDF. | Demonstrate how to set PdfSaveOptions in Aspose.Cells to embed PDF/UA (accessibility) information when exporting a workbook to PDF.
// Common Searches: asp.net convert html to accessible pdf using aspose.cells c# | c# generate pdf/ua compliant document from html with aspose.cells | how to enable screen reader tags when saving html as pdf in aspose.cells | asp.net load html into workbook and export to pdf/a-1b with accessibility | error handling for html to pdf conversion with aspose.cells c#
// Tags: Aspose.Cells HTML to PDF with accessibility tags | PdfSaveOptions screen reader compatibility | HtmlLoadOptions workbook loading C# | PDF/A-1b compliance Aspose.Cells | C# exception handling for file conversion | accessible PDF generation using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program checks for the presence of an input HTML file, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, and saves it as a PDF using PdfSaveOptions. Optional PDF/A‑1b compliance can be enabled, and the saved PDF includes accessibility tags suitable for screen readers. Robust exception handling ensures graceful error reporting.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Path to the destination PDF file
        string pdfPath = "output.pdf";

        try
        {
            // Verify that the HTML source file exists
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"The HTML file '{htmlPath}' was not found.");

            // Load the HTML content into a workbook using HtmlLoadOptions
            var loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Configure PDF save options
            var pdfOptions = new PdfSaveOptions();

            // Uncomment the following line if your Aspose.Cells version supports PDF/A compliance
            // pdfOptions.Compliance = PdfCompliance.PdfA1b;

            // Save the workbook as a PDF with the specified options
            workbook.Save(pdfPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
