// Title: How to convert an Excel file with WordArt gradient fill to PDF/A‑1b while embedding an ICC profile using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx workbook, sets PdfSaveOptions.Compliance to PdfA1b to embed the standard ICC profile, and saves the workbook as a PDF preserving the WordArt gradient fill. | Show a step‑by‑step example of using Aspose.Cells to export an Excel worksheet containing WordArt with gradient shading to a PDF/A‑1b document that includes an embedded ICC color profile.
// Common Searches: Aspose.Cells export Excel WordArt gradient to PDF/A with ICC profile in C# | C# save workbook as PDF/A‑1b preserving gradient fill of WordArt | embed ICC color profile when converting Excel to PDF using Aspose.Cells | PDF/A compliance for WordArt gradient colors Aspose.Cells .NET example
// Tags: Aspose.Cells PDF/A conversion with WordArt gradient | embed ICC profile in PDF generated from Excel | preserve gradient fill during Excel to PDF export | PdfSaveOptions.Compliance PdfA1b C# | convert Excel workbook to PDF/A1b using .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace Example
{
    // The sample checks for the input .xlsx file, loads it with Aspose.Cells, configures PdfSaveOptions.Compliance to PdfA1b (which forces embedding of the standard ICC profile), and saves the workbook as a PDF/A‑1b document, ensuring the WordArt gradient fill is retained.
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
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the Excel workbook that contains WordArt with gradient fill
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // PDF/A compliance forces embedding of an ICC profile, preserving color fidelity
                    Compliance = PdfCompliance.PdfA1b
                    // The EmbedStandardPdfFonts property is not available in this version of Aspose.Cells
                };

                // Save the workbook as PDF; the gradient fill of the WordArt will be retained
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
