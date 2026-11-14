// Title: Export an Excel workbook to PDF with high‑resolution output to keep slicer text sharp using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export Excel to PDF with high DPI in C# | how to preserve slicer label clarity when saving workbook as PDF using Aspose.Cells | C# set PDF image resolution for slicer rendering with Aspose.Cells | PdfSaveOptions OnePagePerSheet and high‑resolution PDF export example | sharp slicer text PDF conversion Aspose.Cells .NET
// Tags: PdfSaveOptions high DPI Aspose.Cells | export Excel to PDF slicer clarity | C# Aspose.Cells PDF rendering resolution | OnePagePerSheet PDF save option | preserve slicer text quality PDF conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, configures PdfSaveOptions (including OnePagePerSheet), and saves the workbook as a PDF. It notes that ImageResolution and SlicerResolution properties are unavailable in the current Aspose.Cells version, so high‑resolution output relies on the library's default rendering.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found – {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Render each sheet on a separate page (optional, improves readability)
                OnePagePerSheet = true
                // ImageResolution and SlicerResolution properties are not available in the current Aspose.Cells version.
                // If needed, they can be set when using a newer version that supports them.
            };

            // Export the workbook to PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
