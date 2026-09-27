// Title: Rasterize an Excel worksheet with embedded SVG graphics to a single PNG file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a .xlsx workbook with Aspose.Cells, selects the first worksheet, and renders it to a PNG image so that any SVG shapes are rasterized. | Show how to set ImageOrPrintOptions (OnePagePerSheet) and use SheetRender to export a worksheet as one PNG page in Aspose.Cells. | Provide robust error‑handling for missing input files and runtime exceptions when converting an Excel sheet containing SVG to PNG in C#.
// Common Searches: aspnet how to convert Excel sheet with SVG objects to PNG using Aspose.Cells | c# rasterize embedded SVG in .xlsx to PNG image programmatically | export first worksheet as single PNG page with Aspose.Cells .NET | handle FileNotFoundException when rendering Excel to PNG with Aspose.Cells | use ImageOrPrintOptions OnePagePerSheet for PNG output from Excel
// Tags: rasterize worksheet to PNG Aspose.Cells | SheetRender export PNG one page | ImageOrPrintOptions OnePagePerSheet .NET | convert embedded SVG to raster image C# | missing workbook file handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads an Excel workbook, selects the first worksheet, and uses Aspose.Cells' SheetRender with ImageOrPrintOptions (OnePagePerSheet) to rasterize the sheet—including any embedded SVG graphics—into a single PNG file. Includes a check for the input file's existence and catches exceptions for graceful error handling.
class SvgToPngRasterizer
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.png";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the SVG image.
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or any specific worksheet by index/name).
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure image rendering options to produce a PNG.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Render the entire sheet on a single page.
                OnePagePerSheet = true
                // The default image format is PNG, so no explicit setting is required.
            };

            // Create a SheetRender object for the worksheet with the specified options.
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);

            // Render the first (and only) page of the sheet to a PNG file.
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Successfully rasterized worksheet to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
