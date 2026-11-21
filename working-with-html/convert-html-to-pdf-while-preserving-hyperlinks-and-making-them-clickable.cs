// Title: Convert an HTML file to PDF while preserving clickable hyperlinks using Aspose.Cells for .NET (C#)
// Common Searches: how to preserve hyperlink functionality when converting HTML to PDF with Aspose.Cells C# | Aspose.Cells C# convert HTML table to PDF with active links | C# example for exporting an HTML workbook to PDF while keeping links clickable using Aspose.Cells | save HTML file as PDF with clickable URLs using Aspose.Cells .NET
// Tags: Aspose.Cells HTML workbook PDF export | hyperlink retention PdfSaveOptions Aspose.Cells | load HTML into Workbook C# Aspose.Cells | clickable links PDF generation Aspose.Cells | PDF/A compliance optional Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions (hyperlinks are retained by default), and saves the workbook as a PDF, producing a PDF where the original HTML hyperlinks remain clickable.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Paths for input HTML and output PDF
        string htmlPath = "input.html";
        string pdfPath = "output.pdf";

        try
        {
            // Verify that the source HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The input file '{htmlPath}' was not found.");
                return;
            }

            // Load the HTML file into a Workbook object.
            // Aspose.Cells can interpret HTML tables and hyperlinks as spreadsheet data.
            Workbook workbook = new Workbook(htmlPath);

            // Configure PDF save options (hyperlinks are preserved by default).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: set PDF/A compliance if required.
                // Compliance = PdfCompliance.PdfA1b
            };

            // Save the workbook as a PDF file with the specified options.
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"Conversion completed. PDF saved to: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
