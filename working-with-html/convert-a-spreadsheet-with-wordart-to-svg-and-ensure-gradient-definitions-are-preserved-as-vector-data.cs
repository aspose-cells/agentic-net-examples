// Title: Export Gradient‑filled WordArt to SVG with Vector Gradients – Aspose.Cells C# Example
// Description: This example creates a workbook, adds a WordArt shape with a preset gradient fill, configures SvgImageOptions to keep the gradient as vector data, and uses ToImage to generate an SVG file that preserves the gradient definitions. The workbook is also saved for Excel verification.
// Keywords: Aspose.Cells | C# | WordArt to SVG | gradient fill SVG | SvgImageOptions | vector gradient export | Excel shape conversion | preserve SVG gradients
// Common Searches: Aspose.Cells export WordArt as SVG | keep gradient fill vector when converting Excel to SVG | C# convert WordArt gradient to SVG | SvgImageOptions gradient settings Aspose | how to preserve gradients in SVG output from Aspose.Cells
// Developer Intent: Generate an SVG file from a gradient‑filled WordArt shape while retaining the gradient as scalable vector data.
// Use Cases: Create web‑ready SVG assets from Excel WordArt that contain gradient fills. | Batch‑export multiple gradient WordArt shapes in a worksheet to individual SVG files. | Produce both an Excel workbook and corresponding SVG graphics for marketing or UI design.
// AI Prompts: Write C# code that scans a worksheet for WordArt shapes with gradient fills and exports each to an SVG using Aspose.Cells, preserving vector gradients. | Explain how FitToViewPort, CssPrefix, and EmbeddedFontType affect the SVG output of gradient‑filled WordArt in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsWordArtToSvg
{
    // This example creates a workbook, adds a WordArt shape with a preset gradient fill, configures SvgImageOptions to keep the gradient as vector data, and uses ToImage to generate an SVG file that preserves the gradient definitions. The workbook is also saved for Excel verification.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the shape collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add a WordArt shape that uses a gradient fill (preset style 7)
            // Parameters: style, text, topRow, top, leftColumn, left, height, width
            Shape wordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7, // Gradient Fill - Blue, Accent 1, Reflection
                "Aspose.Cells",
                2,   // topRow
                0,   // top offset (pixels)
                2,   // leftColumn
                0,   // left offset (pixels)
                100, // height (pixels)
                400  // width (pixels)
            );

            // Configure SVG rendering options
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                // Ensure the generated SVG fits the viewport
                FitToViewPort = true,
                // Optional: set a CSS prefix to avoid naming conflicts
                CssPrefix = "wordart-",
                // Keep the default font embedding (None) – gradients are preserved as vector data
                EmbeddedFontType = SvgEmbeddedFontType.None,
                // Explicitly set the image type to SVG
                ImageType = ImageType.Svg
            };

            // Export the WordArt shape to an SVG file
            // The ToImage method respects the SvgImageOptions and retains gradient definitions
            wordArt.ToImage("WordArtGradient.svg", svgOptions);

            // Optionally, save the workbook to verify the WordArt is present in Excel format
            workbook.Save("WordArtWorkbook.xlsx");
        }
    }
}
