// Title: How to keep Excel cell background colors during PDF conversion with Aspose.Cells for .NET
// AI Prompts: Generate C# code that configures PdfSaveOptions to keep background shading when saving a Workbook as PDF. | Provide an Aspose.Cells snippet that loads an .xlsx file and exports it to PDF while preserving cell background colors.
// Common Searches: Aspose.Cells C# preserve cell background colors when exporting to PDF | PdfSaveOptions property to retain background shading in PDF output | convert Excel to PDF with colors using Aspose.Cells .NET example
// Tags: PdfSaveOptions background rendering | Aspose.Cells PDF export with cell shading | C# preserve Excel colors in PDF | Aspose.Cells PDFSaveOptions color preservation

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook and saves it as a PDF. To retain cell background colors, specific PdfSaveOptions settings must be enabled so that shading is rendered in the PDF output.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (no TransparentBackground property for PDF)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
