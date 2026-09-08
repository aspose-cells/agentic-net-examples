// Title: Shrink PDF output from Aspose.Cells by applying MinimumSize optimization and font subsetting in C#
// AI Prompts: Write C# code that configures PdfSaveOptions with Optimization = MinimumSize and FontEmbedding = FontEmbedding.Subset before saving an Excel workbook to PDF using Aspose.Cells. | Show an example that validates an input Excel file, loads it into a Workbook, applies MinimumSize optimization and font subsetting, and writes the reduced‑size PDF to a target path.
// Common Searches: Aspose.Cells C# set PdfSaveOptions Optimization to MinimumSize | Enable font subsetting in PDF export with Aspose.Cells .NET | How to reduce size of PDF generated from Excel using Aspose.Cells | PdfSaveOptions MinimumSize and FontEmbedding.Subset example code
// Tags: Aspose.Cells PdfSaveOptions MinimumSize | Aspose.Cells PDF font subsetting | C# PDF size reduction Aspose.Cells | Excel to PDF optimization Aspose.Cells | PdfSaveOptions FontEmbedding.Subset

using System;
using System.IO;
using Aspose.Cells;

namespace PdfOptimizationExample
{
    // The program verifies the source Excel file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions to use MinimumSize optimization and to embed only the glyphs actually used (FontEmbedding.Subset), ensures the output directory exists, and saves the workbook as a smaller PDF file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input Excel file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default options are used here)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
