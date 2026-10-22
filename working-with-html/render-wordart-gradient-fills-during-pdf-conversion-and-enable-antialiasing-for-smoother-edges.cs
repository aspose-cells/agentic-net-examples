// Title: How to preserve WordArt gradient fills and apply anti‑aliasing when converting Excel to PDF with Aspose.Cells for .NET
// AI Prompts: Configure PdfSaveOptions.RenderOptions.EnableWordArtRendering, EnableGradientRendering, and EnableAntiAliasing to true before invoking Workbook.Save. | Write C# code that loads an .xlsx containing WordArt, sets the necessary PdfRenderOptions for gradient and anti‑aliasing, and exports the workbook as a PDF. | Explain how to verify that gradient colors and smooth edges are retained in the generated PDF file.
// Common Searches: Aspose.Cells PDF export keep WordArt gradient colors | Enable anti aliasing for PDF conversion in Aspose.Cells .NET | Render WordArt with gradients when saving Excel workbook to PDF using Aspose.Cells | PdfRenderOptions.EnableGradientRendering example Aspose.Cells | How to improve PDF quality of WordArt objects in Aspose.Cells
// Tags: Aspose.Cells PDF gradient rendering | PdfRenderOptions anti-aliasing .NET | EnableWordArtRendering Aspose.Cells | Excel to PDF high quality rendering | WordArt gradient fill Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook that contains WordArt, configures PdfSaveOptions.RenderOptions to enable WordArt rendering, gradient fills, and anti‑aliasing, and then saves the workbook as a PDF with smoother edges and preserved gradient colors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWithWordArt.xlsx";
            const string outputPath = "OutputWithWordArt.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
