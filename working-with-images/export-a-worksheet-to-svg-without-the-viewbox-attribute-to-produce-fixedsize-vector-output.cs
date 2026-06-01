using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.svg";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet to be exported
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure SVG rendering options
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                // Disable viewBox attribute for fixed‑size output
                FitToViewPort = false,
                // Render the whole sheet on a single page
                OnePagePerSheet = true
            };

            // Render the worksheet to an SVG file
            SheetRender sheetRender = new SheetRender(worksheet, svgOptions);
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Worksheet exported to SVG without viewBox: \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}