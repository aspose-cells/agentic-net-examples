// Title: Export an Excel worksheet to PDF while preserving cell background colors and borders with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export worksheet to PDF keep cell background colors | How to retain Excel cell borders when saving as PDF using Aspose.Cells | PdfSaveOptions settings to preserve formatting in PDF conversion .NET | Convert XLSX to PDF with original styling using Aspose.Cells library | C# example for exporting Excel to PDF without losing cell formatting
// Tags: Aspose.Cells PdfSaveOptions preserve cell formatting | C# export Excel worksheet to PDF with background colors | retain borders during Excel to PDF conversion Aspose | Excel to PDF layout options Aspose.Cells .NET | save workbook as PDF maintaining styling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing XLSX file with Aspose.Cells, configures PdfSaveOptions to keep the original layout, and saves the workbook as a PDF. By default the conversion retains cell background colors and borders, producing a PDF that mirrors the worksheet's visual appearance.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the worksheet to export
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep original layout (do not force all columns onto a single page)
                AllColumnsInOnePagePerSheet = false,
                OnePagePerSheet = false
                // RenderGridLines and SaveTransparent are not available in this version;
                // default behavior retains cell borders and background colors.
            };

            // Export the workbook to PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
