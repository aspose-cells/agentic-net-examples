// Title: Aspose.Cells for .NET – Bring a Named Shape (ChartOverlay) to the Front Layer in Excel
// Description: C# example that creates a workbook, adds a rectangle named "ChartOverlay" to the first worksheet, moves the shape to the front using the ToFrontOrBack method, and saves the file as ChartOverlayFront.xlsx. Demonstrates how to control shape Z‑order in Aspose.Cells.
// Keywords: Aspose.Cells C# shape front | ToFrontOrBack method Aspose.Cells | Excel shape Z-order .NET | ChartOverlay shape Aspose | move shape forward Aspose.Cells | C# Excel shape ordering
// Common Searches: Aspose.Cells bring shape to front | C# move Excel shape forward Aspose | How to set Z-order of a shape in Aspose.Cells | ToFrontOrBack example C# | Send rectangle shape to front layer in Excel using Aspose.Cells
// Developer Intent: Place the shape named ChartOverlay above all other worksheet objects by adjusting its Z‑order.
// Use Cases: Ensure a chart overlay remains visible on top of data series in automated financial reports. | Add a watermark that must appear above charts, tables, and images in generated spreadsheets. | Reorder multiple shapes to define visual hierarchy when building dashboards programmatically.
// AI Prompts: Generate C# code with Aspose.Cells that moves a shape called ChartOverlay to the front layer of a worksheet. | Explain the ToFrontOrBack method parameters and how they affect shape ordering in Aspose.Cells. | Show an example that sends a shape to the back layer, then brings it to the front, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, adds a rectangle named "ChartOverlay" to the first worksheet, moves the shape to the front using the ToFrontOrBack method, and saves the file as ChartOverlayFront.xlsx. Demonstrates how to control shape Z‑order in Aspose.Cells.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape and give it the name "ChartOverlay"
        Shape chartOverlay = worksheet.Shapes.AddRectangle(10, 10, 200, 100, 0, 0);
        chartOverlay.Name = "ChartOverlay";

        // Bring the shape to the front (any positive integer moves it forward)
        chartOverlay.ToFrontOrBack(1);

        // Save the workbook
        workbook.Save("ChartOverlayFront.xlsx");
    }
}
