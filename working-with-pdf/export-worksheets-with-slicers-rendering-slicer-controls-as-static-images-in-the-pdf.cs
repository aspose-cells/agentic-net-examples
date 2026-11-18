// Title: Convert an Excel workbook with slicers to PDF, rendering slicers as static images using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel file containing slicers to PDF with slicer images | How to save a workbook with slicer controls as static images in PDF using Aspose.Cells | Aspose.Cells PdfSaveOptions render slicer objects as images in PDF output | Programmatically export Excel worksheet with slicers to PDF in .NET | Save Excel workbook with slicers to PDF without interactive slicer functionality
// Tags: Aspose.Cells PDF export slicer images | PdfSaveOptions static slicer rendering | C# workbook to PDF including slicers | Excel slicer to PDF conversion Aspose | save Excel as PDF static slicer controls

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies the input Excel file, loads it with Aspose.Cells, uses the default PdfSaveOptions (which automatically render slicer controls as static images), and saves the workbook as a PDF while handling potential runtime errors.
class ExportWithSlicers
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that contains slicers
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default rendering already saves slicers as static images)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export the workbook to PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
