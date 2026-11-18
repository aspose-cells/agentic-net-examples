// Title: Embedding used fonts during Excel-to-PDF conversion with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells embed fonts in PDF to avoid missing glyphs | PdfSaveOptions EmbedStandardWindowsFonts true example | How to ensure all fonts are embedded when saving an Excel workbook as PDF using Aspose.Cells | Aspose.Cells PDF conversion missing characters font embedding workaround | Save workbook to PDF with embedded Windows fonts using Aspose.Cells .NET | Prevent missing glyphs in PDF generated from Excel with Aspose.Cells
// Tags: Aspose.Cells PDF font embedding | PdfSaveOptions EmbedStandardWindowsFonts | Excel to PDF conversion .NET | prevent missing glyphs Aspose.Cells | embed used fonts Aspose.Cells PDF | C# Aspose.Cells PDFSaveOptions example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, sets PdfSaveOptions.EmbedStandardWindowsFonts to true, ensures the output folder exists, and saves the workbook as a PDF. It demonstrates how to embed standard Windows fonts during Excel‑to‑PDF conversion with Aspose.Cells for .NET while handling common file‑related errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Embed standard Windows fonts for completeness
                EmbedStandardWindowsFonts = true
                // Note: Aspose.Cells does not expose an EmbedAllFonts property.
                // Standard fonts are embedded automatically when possible.
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
