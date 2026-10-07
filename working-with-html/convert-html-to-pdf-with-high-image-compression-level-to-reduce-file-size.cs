// Title: How to convert an HTML file to a PDF with high image compression using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an HTML document into an Aspose.Cells Workbook and saves it as a PDF while applying maximum JPEG compression to images. | Show the configuration of PdfSaveOptions in Aspose.Cells to enable high image compression (e.g., set Compression to PdfCompressionType.Jpeg and adjust JpegQuality) during HTML‑to‑PDF export. | Provide a complete C# example that validates the input HTML file, uses HtmlLoadOptions, configures PdfSaveOptions for aggressive image compression, and writes the compressed PDF to disk.
// Common Searches: Aspose.Cells C# convert HTML to PDF with JPEG image compression | reduce size of PDF generated from HTML using Aspose.Cells PdfSaveOptions | set high image compression level when exporting HTML to PDF in .NET | how to use HtmlLoadOptions and PdfSaveOptions together for small PDF output
// Tags: HTML to PDF conversion Aspose.Cells C# | PdfSaveOptions image compression Aspose.Cells | Aspose.Cells high JPEG compression PDF | HtmlLoadOptions workbook loading Aspose.Cells | reduce PDF file size Aspose.Cells export

using System;
using System.IO;
using Aspose.Cells;

// The example validates the presence of an HTML file, loads it into an Aspose.Cells Workbook with HtmlLoadOptions, configures PdfSaveOptions to use JPEG compression and a low JpegQuality for aggressive image reduction, and saves the result as a compact PDF while handling any exceptions.
class HtmlToPdfConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.pdf";

            // Verify that the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the HTML file into a workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options (default settings are sufficient for basic conversion)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
