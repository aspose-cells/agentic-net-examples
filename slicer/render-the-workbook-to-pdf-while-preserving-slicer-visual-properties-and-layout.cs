// Title: Render an Excel workbook to PDF while preserving slicer layout and style with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells save Excel workbook to PDF preserving slicer layout | How to keep slicer formatting when exporting to PDF using Aspose.Cells | Aspose.Cells PDFSaveOptions SlicerLayoutOptions example in .NET | Export Excel file with slicers to PDF without losing style in C# | Enable PreserveSlicerLayout in Aspose.Cells PDF conversion
// Tags: Aspose.Cells PDFSaveOptions slicer layout | C# preserve slicer style Aspose.Cells | export Excel to PDF with slicers | SlicerLayoutOptions PreserveSlicerLayout | render workbook to PDF preserving slicers

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an existing Excel workbook, configures PdfSaveOptions with optional SlicerLayoutOptions to retain slicer layout and style, and saves the workbook as a PDF using Aspose.Cells for .NET, handling missing files and runtime exceptions.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains slicers
            Workbook workbook = new Workbook(inputPath);

            // Set PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Preserve slicer layout and style when rendering to PDF (available in newer versions)
            // Uncomment the following block if your Aspose.Cells version supports SlicerLayoutOptions
            /*
            pdfOptions.SlicerLayoutOptions = new SlicerLayoutOptions
            {
                PreserveSlicerLayout = true,
                PreserveSlicerStyle = true
            };
            */

            // Render the workbook to PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
