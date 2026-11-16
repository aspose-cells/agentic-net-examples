// Title: C# Example: Convert an HTML file that uses ::before and ::after CSS pseudo‑elements to PDF with Aspose.Cells
// Common Searches: how to keep CSS ::before and ::after styles when converting HTML to PDF using Aspose.Cells in C# | Aspose.Cells C# load HTML containing pseudo‑elements and export to PDF | C# code to convert an HTML document with CSS pseudo‑elements to PDF via Aspose.Cells | preserve CSS pseudo‑elements during HTML to PDF conversion in .NET using Aspose.Cells | load HTML workbook with LoadOptions Html and save as PDF while retaining layout
// Tags: Aspose.Cells HTML to PDF conversion with pseudo‑elements | C# LoadOptions Html workbook loading | handle CSS pseudo‑elements in PDF export | Aspose.Cells workbook.Save PDF with CSS styling | convert HTML file to PDF using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input HTML file, loads it into an Aspose.Cells Workbook using LoadOptions set to HTML (which parses CSS including ::before and ::after pseudo‑elements), and then saves the workbook as a PDF, preserving the original layout and pseudo‑element styling.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file that may contain CSS pseudo‑elements (::before, ::after)
        string htmlFilePath = "input.html";

        // Desired output PDF file path
        string pdfFilePath = "output.pdf";

        // Verify that the input HTML file exists to avoid FileNotFoundException
        if (!File.Exists(htmlFilePath))
        {
            Console.WriteLine($"Input file not found: {htmlFilePath}");
            return;
        }

        try
        {
            // Load the HTML document into a workbook. Aspose.Cells parses CSS and supports pseudo‑elements.
            var loadOptions = new LoadOptions(LoadFormat.Html);
            var workbook = new Workbook(htmlFilePath, loadOptions);

            // Save the workbook as PDF. Layout is preserved automatically.
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
