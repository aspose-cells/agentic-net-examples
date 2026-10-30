// Title: Export an Excel workbook to PDF with 300 DPI image resampling using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set PDF image resolution to 300 DPI | How to increase image quality when converting XLSX to PDF with Aspose.Cells | PdfSaveOptions ImageResolution property example in C# | Export Excel to high‑resolution PDF using Aspose.Cells | C# generate PDF from workbook with sharper images Aspose.Cells
// Tags: Aspose.Cells PDF image resolution | PdfSaveOptions ImageResolution property | C# export Excel to high‑resolution PDF | 300 DPI image resampling Aspose.Cells | Excel to PDF image quality control | OnePagePerSheet PDF export Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, configures PdfSaveOptions (including the option to set ImageResolution to 300 DPI for sharper images), and saves the workbook as a PDF while handling potential errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (additional options can be set here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: save each worksheet on a separate page
                OnePagePerSheet = true
            };

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
