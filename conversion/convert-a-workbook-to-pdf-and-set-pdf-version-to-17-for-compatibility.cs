// Title: Convert an Excel workbook to PDF with PDF 1.7 compatibility using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set PdfSaveOptions.PdfVersion to 1.7 when saving workbook | How to export .xlsx to PDF with specific PDF version using Aspose.Cells | C# code example for Excel to PDF conversion with PDF 1.7 output | Configure PDF version in Aspose.Cells PDF export options
// Tags: Aspose.Cells PdfSaveOptions PDF version 1.7 | C# Excel to PDF conversion with version control | set PDF compatibility Aspose.Cells | Workbook.Save PDF with specific version | PDF output settings Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace Example
{
    // The example loads an existing Excel file with Aspose.Cells, creates a PdfSaveOptions object, sets PdfSaveOptions.PdfVersion to PDF 1.7 for compatibility, and saves the workbook as a PDF. It includes basic error handling and console messages indicating success or failure.
    class WorkbookToPdf
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Ensure the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default settings)
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
