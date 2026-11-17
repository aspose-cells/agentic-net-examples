// Title: How to set a specific PDF version with PdfSaveOptions when converting Excel to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set PDF version when saving workbook to PDF | PdfSaveOptions PDF/A compliance example in Aspose.Cells .NET | Save Excel as PDF compatible with older PDF readers using Aspose.Cells | Configure PdfCompliance property in Aspose.Cells PdfSaveOptions | Convert .xlsx to PDF with specific PDF version using Aspose.Cells | Aspose.Cells output PDF version 1.4 from Excel workbook
// Tags: PdfSaveOptions PDF version configuration | Aspose.Cells PDF compliance setting | C# convert Excel to PDF specific version | Aspose.Cells older PDF reader compatibility | PdfCompliance PdfA1b Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // The sample verifies the input Excel file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object where the PdfCompliance (e.g., PdfA1b) can be set to target a specific PDF version, and saves the workbook as a PDF while handling any runtime exceptions.
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

                // Load the workbook from the existing Excel file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (no need to set SaveFormat; it's inferred from the options type)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Example: set PDF compliance if required
                    // PdfCompliance = PdfCompliance.PdfA1b
                };

                // Save the workbook as a PDF using the configured options
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
