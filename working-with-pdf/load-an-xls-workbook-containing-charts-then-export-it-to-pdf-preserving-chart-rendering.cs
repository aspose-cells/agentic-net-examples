// Title: Convert an XLS workbook with embedded charts to PDF while preserving chart quality using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xls file containing charts, configures PdfSaveOptions to retain chart rendering as vector graphics, and saves the workbook as a PDF with Aspose.Cells. | Show how to set OnePagePerSheet and related PDF save options in Aspose.Cells to ensure charts are rendered correctly during Excel‑to‑PDF conversion.
// Common Searches: asp.net how to export Excel .xls with charts to PDF using Aspose.Cells preserving vector graphics | c# convert workbook containing charts to PDF with chart rendering intact Aspose.Cells | save Excel file with embedded charts as PDF without losing quality Aspose.Cells .NET
// Tags: Aspose.Cells PDF conversion with chart rendering | C# load XLS workbook and export to PDF | PdfSaveOptions preserve chart vector graphics | Excel to PDF conversion preserving charts .NET | OnePagePerSheet false Aspose.Cells PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks for the presence of an input XLS file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions (including OnePagePerSheet = false) to keep charts rendered as vector graphics, and saves the workbook as a PDF, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the XLS workbook that contains charts
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to preserve chart rendering
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep each sheet on its own page (optional, adjust as needed)
                OnePagePerSheet = false
                // Charts are rendered as vector graphics by default
            };

            // Export the workbook to PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
