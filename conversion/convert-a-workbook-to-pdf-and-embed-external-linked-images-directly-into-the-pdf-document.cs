// Title: Convert an Excel workbook to PDF while embedding external linked images using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert XLSX to PDF with linked images embedded | How to embed external pictures in PDF when saving Excel workbook with Aspose.Cells | Aspose.Cells PdfSaveOptions to include linked images in PDF output | Saving Excel file as PDF with embedded linked images using .NET SDK | Prevent missing images in PDF conversion of Excel using Aspose.Cells
// Tags: Aspose.Cells PDFSaveOptions embed linked images | Excel workbook to PDF conversion C# | include external pictures in PDF Aspose.Cells | save workbook as PDF with images .NET | C# Aspose.Cells linked image handling

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for the input XLSX file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions (which automatically embed any externally linked images), and saves the workbook as a PDF, with basic error handling.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default behavior embeds linked images)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF with the specified options
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
