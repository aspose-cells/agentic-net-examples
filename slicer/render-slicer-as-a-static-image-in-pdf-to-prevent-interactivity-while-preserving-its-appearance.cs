// Title: Convert an Excel slicer to a static image when saving to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel slicer as non‑interactive image in PDF | How to disable slicer interactivity in PDF output using Aspose.Cells | C# PdfSaveOptions OnePagePerSheet slicer rendering static image | Save workbook to PDF while preserving slicer appearance with Aspose.Cells | Render Excel slicer as graphic in PDF using Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | static slicer rendering in PDF | export Excel slicer as image C# | disable slicer interactivity Aspose.Cells | Excel to PDF conversion preserving slicer appearance

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook containing a slicer, configures PdfSaveOptions with OnePagePerSheet set to true, and saves the workbook as a PDF. This renders the slicer as a static graphic, removing interactivity while keeping its visual appearance.
class SlicerToStaticPdf
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing Excel workbook that contains a slicer.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options.
            // OnePagePerSheet ensures that all objects (including slicers) are rendered
            // as static graphics on the PDF page, removing any interactivity.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as a PDF. The slicer will appear as a static image.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
