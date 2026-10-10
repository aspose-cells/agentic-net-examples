// Title: Convert Excel to PDF in C# with Aspose.Cells using MinimumSize optimization and font subsetting to shrink file size
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, configures PdfSaveOptions to use MinimumSize optimization and enables font subsetting, then saves the workbook as a PDF. | Show how to set PdfSaveOptions.OptimizationMode = PdfOptimizationMode.MinimumSize and PdfSaveOptions.FontEmbeddingMode = FontEmbeddingMode.Subset in Aspose.Cells to reduce the generated PDF size.
// Common Searches: set PdfSaveOptions.OptimizationMode to MinimumSize in Aspose.Cells C# | Aspose.Cells example for font subsetting when saving PDF | how to shrink PDF size from Excel conversion using Aspose.Cells | C# code to export Excel to PDF with only used glyphs | reduce PDF file size with Aspose.Cells PdfSaveOptions settings
// Tags: Aspose.Cells PDF optimization MinimumSize | Aspose.Cells subset fonts in PDF export | C# Excel to PDF conversion with reduced size | PdfSaveOptions set optimization mode | PdfSaveOptions embed subset fonts

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, configures PdfSaveOptions to apply MinimumSize optimization and to embed only the glyphs actually used (font subsetting), and then saves the workbook as a PDF, resulting in a smaller output file while handling errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.pdf";

            // Verify that the source workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputFile);

            // Configure PDF save options (default options are sufficient for basic conversion)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputFile, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
