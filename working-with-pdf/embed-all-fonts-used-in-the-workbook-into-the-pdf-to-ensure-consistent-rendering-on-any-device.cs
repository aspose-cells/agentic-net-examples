// Title: Embedding all fonts while saving an Excel workbook as PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file, sets PdfSaveOptions.EmbedStandardWindowsFonts to true, and saves the workbook as a PDF using Aspose.Cells. | Show how to verify the existence of the input Excel file before converting it to a PDF with embedded fonts in a console application. | Provide a complete example that demonstrates font embedding during Excel‑to‑PDF conversion with Aspose.Cells, including error handling.
// Common Searches: Aspose.Cells embed fonts in PDF conversion C# example | how to ensure fonts are embedded when saving Excel to PDF with Aspose.Cells | PdfSaveOptions EmbedStandardWindowsFonts property usage | C# convert .xlsx to PDF with all fonts embedded using Aspose.Cells | prevent missing fonts in PDF generated from Excel with Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions font embedding | C# Excel to PDF conversion with embedded fonts | EmbedStandardWindowsFonts property Aspose.Cells | PDF generation preserving Excel fonts .NET | Workbook.Save PDF with font embedding Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program checks for the input Excel file, loads it with Aspose.Cells, configures PdfSaveOptions to embed standard Windows fonts, and saves the workbook as a PDF, ensuring consistent font rendering across devices.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Embed standard Windows fonts (EmbedAllFonts is not available in older versions)
                    EmbedStandardWindowsFonts = true
                };

                // Save the workbook as a PDF using the specified options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
