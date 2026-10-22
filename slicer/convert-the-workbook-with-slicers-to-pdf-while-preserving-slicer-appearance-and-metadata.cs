// Title: Convert an Excel workbook containing slicers to PDF while preserving slicer appearance using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with slicers, configures PdfSaveOptions to keep slicer graphics, and saves the workbook as a PDF with Aspose.Cells. | Show how to verify the input Excel file exists, handle exceptions, and log conversion results when exporting slicer‑enabled worksheets to PDF in a .NET console app. | Demonstrate setting PdfSaveOptions properties (e.g., OnePagePerSheet, EmbedStandardFonts) to ensure slicer metadata is retained during Excel‑to‑PDF conversion with Aspose.Cells.
// Common Searches: How to keep slicer formatting when converting Excel to PDF with Aspose.Cells C# | Aspose.Cells PDF export slicer visibility .NET example | C# code to export workbook that has slicers to PDF preserving slicer graphics
// Tags: Aspose.Cells PdfSaveOptions slicer rendering | C# Excel to PDF conversion with slicer preservation | slicer graphics retention Aspose.Cells | Excel workbook slicer export PDF .NET | PDF export preserving slicer metadata Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The sample checks for 'input.xlsx', loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions instance, and saves the workbook as 'output.pdf' while retaining slicer graphics and handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
