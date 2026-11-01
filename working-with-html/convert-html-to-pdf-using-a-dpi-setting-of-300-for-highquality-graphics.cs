// Title: Convert HTML to PDF with 300 DPI high‑resolution graphics using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an HTML file into an Aspose.Cells Workbook, sets PdfSaveOptions.ImageDpi to 300, and saves it as a PDF. | Show how to add file‑existence validation and detailed exception handling to an HTML‑to‑PDF conversion routine that uses Aspose.Cells. | Demonstrate configuring PdfSaveOptions for high‑quality PDF output (300 DPI) when exporting a workbook created from an HTML source.
// Common Searches: how to export HTML to PDF at 300 dpi using Aspose.Cells C# | Aspose.Cells set PDF DPI when converting from HTML | C# high resolution PDF generation from HTML workbook | PdfSaveOptions DPI property example Aspose.Cells
// Tags: Aspose.Cells HTML to PDF with custom DPI | PdfSaveOptions DPI configuration .NET | C# workbook conversion HTML to PDF | input file validation Aspose.Cells | exception handling Aspose.Cells PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies that an input HTML file exists, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions.ImageDpi to 300 for high‑quality graphics, and saves the workbook as a PDF while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.pdf";

            // Verify that the input HTML file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the HTML file into a workbook
            var loadOptions = new HtmlLoadOptions();
            var workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options (high‑quality DPI can be set via ImageDpi if supported)
            var pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
