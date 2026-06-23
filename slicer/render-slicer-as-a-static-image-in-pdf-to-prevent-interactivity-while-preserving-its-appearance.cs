using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class RenderSlicerToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWithSlicer.xlsx";
            const string outputPath = "OutputWithStaticSlicer.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook that contains the slicer
            Workbook workbook = new Workbook(inputPath);

            // Assume the slicer is on the first worksheet and is the first slicer object
            Worksheet sheet = workbook.Worksheets[0];
            SlicerCollection slicers = sheet.Slicers;
            if (slicers.Count > 0)
            {
                // Get the first slicer and make it printable (static image in output)
                Slicer slicer = slicers[0];
                slicer.IsPrintable = true; // Marks slicer as printable; interactivity will be removed
            }

            // Prepare PDF save options (optional, can customize further)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF; the slicer will appear as a static image
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved to PDF with slicer rendered as a static image: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}