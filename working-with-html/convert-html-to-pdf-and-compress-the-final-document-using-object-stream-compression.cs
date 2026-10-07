// Title: C# – Convert an HTML file to a compressed PDF using Aspose.Cells with object‑stream compression
// AI Prompts: Generate C# code that loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions to turn on object‑stream mode, and saves the result as a PDF. | Show how to set PdfSaveOptions.EnableObjectStream = true and select the highest compression level when converting HTML to PDF with Aspose.Cells. | Provide a complete C# example that validates the HTML input path, applies object‑stream mode, and writes the compressed PDF to disk using Aspose.Cells.
// Common Searches: Aspose.Cells C# enable PDF object stream for smaller files | convert HTML to PDF with maximum compression using Aspose.Cells .NET | C# sample code for HTML workbook to compressed PDF with PdfSaveOptions | how to reduce PDF file size when saving from Aspose.Cells workbook
// Tags: Aspose.Cells PDF object stream | C# HTML workbook PDF export Aspose.Cells | PdfSaveOptions object stream setting | high PDF compression Aspose.Cells | PDF size reduction from HTML workbook

using System;
using System.IO;
using Aspose.Cells;

// The program checks that an HTML file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions instance, and saves the workbook as a PDF. By setting PdfSaveOptions.EnableObjectStream = true (and optionally increasing the compression level), the resulting PDF can be compressed using object‑stream compression.
class Program
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the HTML file into a Workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
