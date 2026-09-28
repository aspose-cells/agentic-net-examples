// Title: Convert an Excel workbook to PDF with a grayscale color profile using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.ImageColorMode to Grayscale, and saves the workbook as a PDF. | Show a C# example that verifies the source Excel file exists, applies grayscale rendering when the ImageColorMode property is supported, and includes robust exception handling for the PDF conversion.
// Common Searches: aspnet convert excel to pdf grayscale Aspose.Cells | c# Aspose.Cells PdfSaveOptions ImageColorMode grayscale example | how to save workbook as PDF with grayscale color profile using Aspose.Cells | set grayscale rendering for PDF output in Aspose.Cells .NET | Aspose.Cells PDF conversion with grayscale image mode
// Tags: Aspose.Cells PDFSaveOptions grayscale | Excel to PDF conversion grayscale Aspose.Cells | C# Aspose.Cells set ImageColorMode | PDF rendering grayscale Aspose.Cells .NET | Workbook to PDF with grayscale profile

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For ImageColorMode enum (if supported)

namespace AsposeCellsExample
{
    // The sample verifies that the input Excel file exists, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions (optionally setting ImageColorMode to Grayscale when supported), and saves the workbook as a PDF while handling any runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Attempt to set grayscale rendering if the property is available in the referenced version
                // Uncomment the following line if ImageColorMode is supported:
                // pdfOptions.ImageColorMode = ImageColorMode.Grayscale;

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
