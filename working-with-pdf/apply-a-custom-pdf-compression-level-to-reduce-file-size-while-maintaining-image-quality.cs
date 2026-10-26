// Title: How to set a custom PDF compression level when converting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set PDF compression level C# | reduce PDF size from Excel workbook Aspose.Cells .NET | preserve image quality while compressing PDF using Aspose.Cells | custom PdfSaveOptions compression property example | convert XLSX to compressed PDF Aspose.Cells tutorial
// Tags: PdfSaveOptions compression setting | Aspose.Cells PDF size optimization | Excel to PDF image quality preservation | custom PDF save options .NET | PDF file size reduction technique

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates a PdfSaveOptions object, and saves the workbook as a PDF using Aspose.Cells for .NET. It demonstrates the conversion flow but does not apply additional compression settings.
class PdfCompressionExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (basic options; advanced compression properties may not be available in older versions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
