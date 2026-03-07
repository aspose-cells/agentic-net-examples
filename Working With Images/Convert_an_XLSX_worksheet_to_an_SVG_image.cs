using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class ConvertWorksheetToSvg
{
    static void Main()
    {
        // Load the source Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or any specific worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions();
        svgOptions.ImageType = ImageType.Svg;          // Set output format to SVG
        svgOptions.FitToViewPort = true;              // Make SVG fit the viewport
        svgOptions.OnePagePerSheet = true;            // Render the whole sheet on a single page

        // Create a SheetRender instance with the worksheet and SVG options
        SheetRender renderer = new SheetRender(worksheet, svgOptions);

        // Render the first (and only) page of the worksheet to an SVG file
        renderer.ToImage(0, "output.svg");
    }
}