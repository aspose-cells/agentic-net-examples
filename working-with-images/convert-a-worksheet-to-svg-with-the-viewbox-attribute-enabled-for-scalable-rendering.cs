using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet (creation rule)
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data (optional, demonstrates rendering)
        worksheet.Cells["A1"].PutValue("Item");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["B3"].PutValue(15);

        // Configure SVG rendering options and enable FitToViewPort (property rule)
        SvgImageOptions svgOptions = new SvgImageOptions();
        svgOptions.FitToViewPort = true;                     // Enable viewBox scaling
        svgOptions.ImageType = ImageType.Svg;                // Ensure output format is SVG

        // Render the worksheet to SVG (save rule)
        SheetRender renderer = new SheetRender(worksheet, svgOptions);
        renderer.ToImage(0, "output_fit_to_viewport.svg");

        Console.WriteLine("SVG file generated with FitToViewPort enabled.");
    }
}