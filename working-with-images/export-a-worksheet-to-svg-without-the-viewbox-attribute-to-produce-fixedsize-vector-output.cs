// Title: Export Worksheet to Fixed‑Size SVG (No viewBox) with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, fill cells, and use Aspose.Cells SvgImageOptions with FitToViewPort set to false. The SheetRender class then saves the first worksheet as an SVG file that omits the viewBox attribute, yielding a fixed‑size vector image.
// Keywords: Aspose.Cells | C# | .NET | SVG export | FitToViewPort false | no viewBox | fixed size SVG | SheetRender | Excel to SVG | image rendering options
// Common Searches: Aspose.Cells export worksheet to SVG without viewBox | C# generate fixed size SVG from Excel sheet | disable viewBox in SVG output Aspose.Cells | FitToViewPort false Aspose.Cells example | render Excel worksheet as static SVG .NET
// Developer Intent: Create an SVG file from a worksheet that excludes the viewBox attribute, producing a vector image with predetermined dimensions.
// Use Cases: Embedding a non‑scalable SVG chart in a web page where exact dimensions are required. | Generating printable SVG diagrams from Excel data for consistent layout in reports. | Exporting dashboard sheets to SVG for inclusion in PDFs without responsive scaling.
// AI Prompts: Show how to export all worksheets in a workbook to separate fixed‑size SVG files using Aspose.Cells. | Explain how to specify custom width and height for the SVG when FitToViewPort is disabled. | Provide code to convert the generated fixed‑size SVG to PNG while preserving its dimensions.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, fill cells, and use Aspose.Cells SvgImageOptions with FitToViewPort set to false. The SheetRender class then saves the first worksheet as an SVG file that omits the viewBox attribute, yielding a fixed‑size vector image.
class ExportWorksheetToSvgFixedSize
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Item");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["B3"].PutValue(15);

        // Configure SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions();
        svgOptions.ImageType = ImageType.Svg;      // Ensure SVG output
        svgOptions.FitToViewPort = false;         // Disable viewBox generation for fixed‑size output
        svgOptions.OnePagePerSheet = true;        // Render the whole sheet on a single page (optional)

        // Render the worksheet to an SVG file
        SheetRender renderer = new SheetRender(worksheet, svgOptions);
        renderer.ToImage(0, "output_fixed.svg");

        Console.WriteLine("Worksheet exported to SVG without viewBox.");
    }
}
