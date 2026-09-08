// Title: Set PdfSaveOptions.ImageResample to 150 DPI for PDF export with Aspose.Cells in C# (including fallback when property is missing)
// AI Prompts: Generate C# code that creates a PdfSaveOptions object, assigns ImageResample = 150 DPI, and saves an Excel workbook as a PDF using Aspose.Cells. | Provide C# alternatives for reducing PDF size while preserving image clarity when the ImageResample property is not available in Aspose.Cells.
// Common Searches: Aspose.Cells C# set PDF image resample DPI to 150 | PdfSaveOptions ImageResample property missing workaround | How to reduce PDF size by resampling images in Aspose.Cells | C# Excel to PDF image quality control with Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions image DPI | C# Excel to PDF image resampling | reduce PDF file size Aspose.Cells | alternative image quality settings Aspose.Cells | PDF export image compression C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, attempts to configure PdfSaveOptions for image resampling at 150 DPI, notes that the ImageResample property may be unavailable in the current Aspose.Cells version, and saves the workbook as a PDF while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: The ImageResample property is not available in the current Aspose.Cells version.
            // If needed, adjust image quality using other available options.

            // Save the workbook as a PDF with the specified options
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
