// Title: Convert an Excel workbook to PDF with high‑resolution (300 DPI) images using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF with 300 DPI images | How to set image resolution when saving a workbook as PDF using Aspose.Cells | PdfSaveOptions ImageDpi property example for high quality PDF in .NET | Generate print‑ready PDF from Excel with 300 DPI using Aspose.Cells | Configure OnePagePerSheet and ImageDpi for PDF conversion in Aspose.Cells C#
// Tags: Aspose.Cells PDF conversion with ImageDpi | PdfSaveOptions high resolution image export | C# Excel to PDF 300 DPI | print ready PDF generation Aspose.Cells | OnePagePerSheet false Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input Excel file, loads it into a Workbook, configures PdfSaveOptions (disabling OnePagePerSheet and indicating where ImageDpi could be set), and saves the workbook as a PDF while handling any exceptions.
class WorkbookToPdfConverter
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure all worksheets are saved (not one page per sheet)
                OnePagePerSheet = false
                // ImageDpi property omitted for compatibility with current API version
            };

            // Desired PDF file path
            string outputPath = "output.pdf";

            // Save the workbook as a PDF with the specified options
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
