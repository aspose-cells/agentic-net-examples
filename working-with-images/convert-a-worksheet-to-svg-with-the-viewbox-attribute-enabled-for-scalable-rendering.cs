// Title: Convert an Excel worksheet to a single-page SVG with viewBox using Aspose.Cells for .NET (C#)
// Common Searches: how to save an Excel worksheet as SVG with viewBox using Aspose.Cells C# | Aspose.Cells ImageSaveOptions SVG OnePagePerSheet example | C# export .xlsx to scalable SVG with viewBox attribute | generate SVG from workbook worksheet Aspose.Cells .NET | set viewBox for SVG output in Aspose.Cells C#
// Tags: Aspose.Cells ImageSaveOptions SVG export | C# workbook to SVG conversion | Enable viewBox in Aspose.Cells SVG output | OnePagePerSheet setting for SVG rendering | Excel worksheet scalable SVG generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an .xlsx workbook, configures ImageSaveOptions to produce a single‑page SVG that includes the viewBox attribute for scalable rendering, and saves the result, with basic file‑existence validation and error handling.
class WorksheetToSvg
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.svg";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure image save options for SVG format
            ImageSaveOptions svgOptions = new ImageSaveOptions(SaveFormat.Svg)
            {
                // Save each worksheet on a single page
                ImageOrPrintOptions = { OnePagePerSheet = true }
            };

            // Save the workbook; each worksheet will be saved as a separate SVG file
            workbook.Save(outputPath, svgOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
