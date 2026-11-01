// Title: Convert an Excel worksheet to a scalable SVG with viewBox using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export worksheet to SVG with viewBox attribute | Generate single-page SVG from an Excel sheet using Aspose.Cells | C# code example for converting .xlsx to .svg with Aspose.Cells | Enable viewBox when saving Excel as SVG in .NET | How to render an Excel worksheet as scalable SVG using Aspose.Cells
// Tags: Aspose.Cells worksheet SVG export | ImageOrPrintOptions SaveFormat Svg setting | OnePagePerSheet option for SVG output | SheetRender ToImage method for SVG | viewBox enabled SVG rendering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample verifies the input .xlsx file, loads it into a Workbook, selects the first worksheet, configures ImageOrPrintOptions with SaveFormat.Svg and OnePagePerSheet=true to enable viewBox, and uses SheetRender.ToImage to generate a single‑page, scalable SVG file, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.svg";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or modify the index/name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure rendering options for SVG output
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                // Set the output format to SVG
                SaveFormat = SaveFormat.Svg,

                // Render the entire sheet on a single page (required for SVG)
                OnePagePerSheet = true
            };

            // Create a SheetRender object for the worksheet with the specified options
            SheetRender sheetRender = new SheetRender(worksheet, renderOptions);

            // Render the first page of the worksheet to an SVG file
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Worksheet rendered successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
