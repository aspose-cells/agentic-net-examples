// Title: Maintain conditional formatting colors while converting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook to PDF without losing conditional formatting colors | How to export Excel to PDF preserving cell color rules using PdfSaveOptions | C# code example for keeping conditional formatting when converting XLSX to PDF with Aspose.Cells | PDF conversion from Excel with conditional formatting intact in .NET | Aspose.Cells PDF rendering options to retain conditional formatting
// Tags: Aspose.Cells PdfSaveOptions conditional formatting | Excel to PDF conversion preserving colors .NET | C# conditional formatting rendering in PDF | PDF export with cell color rules Aspose.Cells | maintain formatting during Excel PDF save

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example loads an existing XLSX file using Aspose.Cells, creates a PdfSaveOptions object (which renders conditional formatting by default), and saves the workbook as a PDF while handling missing file and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook from the existing Excel file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (conditional formatting is rendered by default)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF with the specified options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Successfully saved PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
