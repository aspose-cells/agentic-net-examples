// Title: How to set PdfSaveOptions.ImageResample to 150 DPI in Aspose.Cells for .NET to shrink PDF size while keeping image quality
// Common Searches: Aspose.Cells C# set ImageResample DPI to 150 for PDF export | reduce PDF file size with image resampling in Aspose.Cells .NET | configure PdfSaveOptions.ImageResample property for high‑quality PDF conversion | adjust image resolution when saving Excel as PDF using Aspose.Cells | best DPI setting for image resampling in Aspose.Cells PDF output
// Tags: PdfSaveOptions image resample DPI | Aspose.Cells PDF image resolution control | optimize PDF size with image resampling | maintain image quality during Excel to PDF conversion | C# configure PdfSaveOptions for PDF output

using System;
using System.IO;
using Aspose.Cells;

// Load the workbook, create a PdfSaveOptions instance, set pdfOptions.ImageResample = 150, then save the workbook as PDF. This resamples embedded images to 150 DPI, reducing the PDF file size while preserving visual clarity.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            var workbook = new Workbook(inputPath);

            // Create PDF save options (default settings)
            var pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
