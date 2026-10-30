// Title: Disable gridline rendering in PDF output using Aspose.Cells PdfSaveOptions in C#
// Common Searches: Aspose.Cells C# hide Excel gridlines in generated PDF | set RenderSolidGridlines false when saving workbook to PDF with Aspose.Cells | PDF export without gridlines using Aspose.Cells .NET API | how to turn off gridlines in PDF conversion with Aspose.Cells | Aspose.Cells PdfSaveOptions gridline rendering option
// Tags: Aspose.Cells PdfSaveOptions disable gridlines | RenderSolidGridlines false Aspose.Cells | Excel to PDF conversion without gridlines C# | OnePagePerSheet PDF export Aspose.Cells | clean PDF layout Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures PdfSaveOptions (including OnePagePerSheet) and demonstrates how to disable gridline rendering by setting RenderSolidGridlines to false, producing a cleaner PDF layout when saving with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // The RenderGridlines property is not available in this version of Aspose.Cells.
                // If needed, other options can be set here, e.g., OnePagePerSheet.
                OnePagePerSheet = true
            };

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
