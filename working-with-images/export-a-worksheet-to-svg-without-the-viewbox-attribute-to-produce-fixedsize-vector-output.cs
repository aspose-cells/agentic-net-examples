// Title: Export Excel Worksheet to Fixed‑Size SVG without viewBox using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, fill cells, configure SvgImageOptions with FitToViewPort = false and OnePagePerSheet = true, and render the first sheet as a fixed‑size SVG file that omits the viewBox attribute.
// Keywords: Aspose.Cells SVG export | C# fixed size SVG | remove viewBox Aspose | SvgImageOptions FitToViewPort false | SheetRender export SVG | one page per sheet SVG | Aspose.Cells .NET image rendering
// Common Searches: Aspose.Cells export worksheet to SVG without viewBox | C# generate fixed size SVG from Excel | disable viewBox in Aspose.Cells SVG output | render whole sheet as single SVG page .NET | FitToViewPort false Aspose.Cells example
// Developer Intent: Create a vector SVG file from an Excel worksheet that has a predetermined canvas size and does not contain a viewBox attribute.
// Use Cases: Embedding a non‑scalable table graphic in PDFs or reports. | Providing consistent‑size SVG icons for dashboards that must not auto‑scale. | Including a fixed‑dimension SVG chart in email templates where responsive scaling is undesirable.
// AI Prompts: Show how to set explicit width and height for the SVG while keeping FitToViewPort disabled. | Generate code that exports each worksheet in a workbook to separate fixed‑size SVG files. | Explain the steps to embed the produced FixedSizeOutput.svg into HTML without automatic scaling.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook, fill cells, configure SvgImageOptions with FitToViewPort = false and OnePagePerSheet = true, and render the first sheet as a fixed‑size SVG file that omits the viewBox attribute.
class ExportWorksheetToFixedSizeSvg
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates an empty workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(85);

            // Configure SVG rendering options
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                // Ensure the SVG does NOT fit to the viewport (no viewBox attribute)
                FitToViewPort = false,
                // Render the whole sheet on a single page
                OnePagePerSheet = true
                // No need to set ImageFormat; SvgImageOptions is specific to SVG output
            };

            // Render the first (and only) page of the worksheet to an SVG file
            SheetRender renderer = new SheetRender(sheet, svgOptions);
            renderer.ToImage(0, "FixedSizeOutput.svg");

            Console.WriteLine("Worksheet exported to FixedSizeOutput.svg without viewBox.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
