// Title: How to embed standard Windows fonts while converting an Excel workbook to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert Excel to PDF with embedded Windows fonts | set EmbedStandardWindowsFonts true in PdfSaveOptions Aspose.Cells | save workbook as PDF with font embedding using Aspose.Cells .NET | C# code to ensure PDF fonts are embedded when exporting Excel with Aspose.Cells | prevent missing fonts in PDF generated from Excel using Aspose.Cells
// Tags: PdfSaveOptions embed standard windows fonts | Aspose.Cells export Excel to PDF with font embedding | C# embed fonts in PDF using Aspose.Cells | cross-platform PDF compatibility Aspose.Cells | Excel to PDF conversion with embedded fonts .NET

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the source Excel file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions to embed standard Windows fonts, creates the output directory if needed, and saves the workbook as a PDF with the specified font‑embedding settings, handling any runtime exceptions.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure standard Windows fonts are embedded
                EmbedStandardWindowsFonts = true
                // Note: FontEmbeddingMode property is not available in this version of Aspose.Cells
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
