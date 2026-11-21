// Title: Convert HTML to PDF using Aspose.Cells for .NET (C#)
// Common Searches: how to load an HTML file into Aspose.Cells workbook in C# | Aspose.Cells PDFSaveOptions to control image quality when saving PDF | reduce PDF file size when converting HTML with Aspose.Cells C# | C# example for converting HTML to PDF with Aspose.Cells | set high image compression level in Aspose.Cells PDF export
// Tags: Aspose.Cells HTML to PDF conversion | PdfSaveOptions image compression control | Aspose.Cells PDF size optimization | load HTML workbook Aspose.Cells | C# Aspose.Cells PDF export example

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for an input HTML file, loads it into an Aspose.Cells Workbook using HtmlLoadOptions, creates a PdfSaveOptions object with default settings, and saves the workbook as a PDF while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.pdf";

            // Verify that the input HTML file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the HTML file into a workbook
            var loadOptions = new HtmlLoadOptions();
            var workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options (using defaults)
            var pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
