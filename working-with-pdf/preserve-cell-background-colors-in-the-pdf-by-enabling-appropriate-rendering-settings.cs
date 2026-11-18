// Title: How to retain Excel cell background colors when converting to PDF with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells keep cell background colors when exporting workbook to PDF | retain Excel cell shading in PDF using Aspose.Cells PdfSaveOptions | preserve fill colors during Excel to PDF conversion with Aspose.Cells .NET | Aspose.Cells PDF conversion without losing cell background color C# example | how to enable background color rendering in PdfSaveOptions for PDF output
// Tags: Aspose.Cells PDF background color preservation | PdfSaveOptions background rendering | Excel to PDF conversion color fidelity Aspose.Cells | C# preserve cell fill when saving as PDF | save workbook as PDF with formatting retained

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This C# example loads an Excel workbook with Aspose.Cells, creates a PdfSaveOptions instance (which by default retains cell background colors), and saves the workbook as a PDF, ensuring that the original cell shading and fill colors are preserved in the PDF output.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default settings preserve cell background colors)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
