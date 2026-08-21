// Title: Export an Aspose.Cells Worksheet to Scalable SVG with viewBox (FitToViewPort) in C#
// Description: Shows how to build a workbook, populate cells, enable SvgImageOptions.FitToViewPort, and use SheetRender to create an SVG file that contains a viewBox attribute for responsive scaling.
// Keywords: Aspose.Cells | C# SVG export | FitToViewPort | viewBox | SheetRender | SvgImageOptions | Excel to SVG | scalable vector graphics | export worksheet as SVG | Aspose.Cells rendering
// Common Searches: Aspose.Cells enable viewBox when exporting to SVG | C# convert Excel worksheet to scalable SVG | FitToViewPort property SVG Aspose.Cells example | How to render worksheet as SVG with viewBox attribute | SheetRender SVG options C#
// Developer Intent: Generate an SVG file from a worksheet that includes a viewBox for flexible rendering on any screen size.
// Use Cases: Create responsive web graphics from Excel data without loss of quality. | Embed worksheet visuals in HTML pages as vector images that adapt to container dimensions. | Produce print‑ready vector files from spreadsheets while preserving layout and scaling.
// AI Prompts: Write C# code using Aspose.Cells to export a worksheet to SVG with FitToViewPort enabled and explain each step. | Describe how the FitToViewPort setting adds a viewBox to the SVG output and why it matters for web scaling. | Provide a sample that converts all worksheets in a workbook to separate SVG files, each with the viewBox attribute.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to build a workbook, populate cells, enable SvgImageOptions.FitToViewPort, and use SheetRender to create an SVG file that contains a viewBox attribute for responsive scaling.
class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(150);
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["B3"].PutValue(250);

        // Configure SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions();
        svgOptions.FitToViewPort = true;                     // Enable viewBox attribute for scalable rendering
        svgOptions.ImageType = Aspose.Cells.Drawing.ImageType.Svg; // Ensure output format is SVG

        // Create a SheetRender instance with the worksheet and SVG options (loading rule)
        SheetRender renderer = new SheetRender(sheet, svgOptions);

        // Render the first page of the worksheet to an SVG file (saving rule)
        renderer.ToImage(0, "worksheet_output.svg");

        Console.WriteLine("Worksheet has been converted to SVG with FitToViewPort enabled.");
    }
}
