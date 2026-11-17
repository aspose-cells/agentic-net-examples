// Title: Disable external resource loading while converting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to turn off external image loading when exporting to PDF | prevent linked files from being fetched during Excel to PDF conversion using Aspose.Cells | improve PDF export speed by disabling external resources in Aspose.Cells PdfSaveOptions | C# save workbook as PDF without loading external content with Aspose.Cells | default PdfSaveOptions behavior for external resources in Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions disable external resources | Excel to PDF conversion performance tuning Aspose.Cells | C# workbook save as PDF without external links | PDF rendering without external images using Aspose.Cells | disable external content loading Aspose.Cells PDF export

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The example verifies the input XLSX file, loads it into a Workbook, creates a PdfSaveOptions instance (which disables external resource loading by default), and saves the workbook as a PDF, reducing rendering time and avoiding external image or link retrieval.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (external resources loading is disabled by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
