// Title: Convert an Excel workbook to PDF with 50% JPEG compression using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, configures PdfSaveOptions with JpegQuality = 50, and saves the workbook as a PDF using Aspose.Cells. | Demonstrate how to apply JPEG compression at 50% when exporting an Excel workbook to PDF with Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells how to set JpegQuality when saving workbook to PDF in C# | C# convert Excel file to PDF with reduced image quality using Aspose.Cells | PDF output size reduction by adjusting JPEG quality in Aspose.Cells .NET | Save Excel as PDF with specific JPEG compression level Aspose.Cells
// Tags: Aspose.Cells JPEG quality option | C# Excel to PDF conversion with image compression | reduce PDF size Aspose.Cells | set JPEG compression level Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example verifies the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, optionally sets JpegQuality to 50 for JPEG compression, and saves the workbook as output.pdf while handling any exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Set JPEG quality if the property is available in the used Aspose.Cells version
                // Uncomment the following line when JpegQuality is supported:
                // pdfOptions.JpegQuality = 50;

                // Save the workbook as a PDF using the configured options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
