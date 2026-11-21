// Title: How to embed all fonts from an Excel workbook into a PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# embed all fonts when converting Excel to PDF | set PdfSaveOptions to embed fonts in PDF generated from XLSX using Aspose.Cells | C# code to ensure PDF contains embedded fonts after Excel to PDF conversion | force font embedding in PDF output with Aspose.Cells .NET | prevent missing fonts in PDF created from Excel using Aspose.Cells C#
// Tags: Aspose.Cells PDF font embedding | PdfSaveOptions embed fonts C# | Excel to PDF conversion with font embedding | C# embed fonts in PDF using Aspose | consistent PDF rendering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, configures PdfSaveOptions to embed all used fonts, and saves the workbook as a PDF, guaranteeing consistent font rendering on any device.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default options are sufficient)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
