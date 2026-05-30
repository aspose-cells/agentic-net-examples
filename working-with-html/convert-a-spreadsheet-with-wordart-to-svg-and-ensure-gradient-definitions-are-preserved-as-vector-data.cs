using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset style that uses gradient fill (WordArtStyle7)
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            2, 0,    // topRow, top offset (pixels)
            2, 0,    // leftColumn, left offset (pixels)
            100, 400 // height, width (pixels)
        );

        // Configure SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions
        {
            ImageType = ImageType.Svg,   // Ensure output format is SVG
            FitToViewPort = true         // Make SVG fit the viewport
            // Gradient fills are preserved automatically as vector data
        };

        // Render the worksheet (including the WordArt) to an SVG file
        SheetRender renderer = new SheetRender(worksheet, svgOptions);
        renderer.ToImage(0, "WordArtGradient.svg");
    }
}