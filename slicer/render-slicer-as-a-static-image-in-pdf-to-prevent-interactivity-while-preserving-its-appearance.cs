using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Namespace for slicer objects

class RenderSlicerAsStaticPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWithSlicer.xlsx";
            const string outputPath = "OutputStaticSlicer.pdf";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing slicers
            Workbook workbook = new Workbook(inputPath);

            // Make each slicer printable so it appears in the PDF as a static image
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Slicer slicer in sheet.Slicers)
                {
                    slicer.IsPrintable = true;
                }
            }

            // Save the workbook as PDF; slicers will be rendered as static images
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved to PDF with slicers rendered as static images: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}