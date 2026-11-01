// Title: Convert an Excel workbook containing slicers to PDF while keeping slicer appearance and layout using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, sets PdfSaveOptions to retain slicer visuals, and saves the workbook as a PDF with Aspose.Cells. | Show how to configure Aspose.Cells PdfSaveOptions so that slicer formatting is not altered when exporting an Excel sheet to PDF in a .NET application.
// Common Searches: Aspose.Cells export Excel to PDF keep slicer formatting | C# convert workbook with slicers to PDF preserving layout | PdfSaveOptions settings to retain slicer visuals in Aspose.Cells | how to prevent slicer distortion when saving Excel as PDF using Aspose.Cells .NET
// Tags: export workbook with slicers to PDF Aspose.Cells | PdfSaveOptions preserve slicer layout | C# Aspose.Cells slicer visual retention | convert Excel slicer to PDF .NET

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the source Excel file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions with OnePagePerSheet and AllColumnsInOnePagePerSheet set to false to maintain the original sheet layout, and saves the workbook as a PDF while preserving slicer appearance. Errors are caught and reported.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not force the sheet onto a single page; retain original layout
                OnePagePerSheet = false,
                AllColumnsInOnePagePerSheet = false
                // Note: Slicer layout preservation is handled automatically in recent versions.
            };

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
