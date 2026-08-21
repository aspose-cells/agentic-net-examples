// Title: Aspose.Cells C# – Send Named Shape “ChartOverlay” to Front Layer (Z‑Order)
// Description: Creates a workbook, adds a column chart, inserts a rectangle shape named ChartOverlay, and uses the ToFrontOrBack method to bring the shape to the front layer so it appears above all other objects before saving as ChartOverlayToFront.xlsx.
// Keywords: Aspose.Cells | C# | shape front layer | ToFrontOrBack | ChartOverlay | Excel shape Z-order | move shape to front | Aspose.Cells .NET | worksheet shapes | chart overlay shape
// Common Searches: Aspose.Cells move shape to front layer C# | How to bring a named shape to front in Excel using Aspose.Cells | ToFrontOrBack method example Aspose.Cells .NET | Set Z‑order for shapes in Aspose.Cells workbook | Chart overlay shape front Aspose.Cells C#
// Developer Intent: Place the shape named ChartOverlay on the front layer so it renders above every chart, image, or other worksheet object.
// Use Cases: Overlay a rectangle with annotations on a chart without obscuring the chart data. | Add a persistent watermark that stays visible over all worksheet content. | Prioritize a specific shape in a multi‑layer report generated programmatically.
// AI Prompts: Generate C# code with Aspose.Cells that moves a shape called ChartOverlay to the front layer of a worksheet. | Create a reusable method that accepts a Worksheet and a shape name, then calls ToFrontOrBack to set the shape to the front. | Provide a try‑catch example that brings a named shape to the front and logs any exceptions in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart, inserts a rectangle shape named ChartOverlay, and uses the ToFrontOrBack method to bring the shape to the front layer so it appears above all other objects before saving as ChartOverlayToFront.xlsx.
class ChartOverlayToFront
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a sample chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 5, 15, 15);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add a rectangle shape that will act as the overlay and give it a name
            // Parameters: type, upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
            Shape chartOverlay = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 6, 6, 0, 0, 200, 200);
            chartOverlay.Name = "ChartOverlay";

            // Bring the shape to the front layer (1 = front, 0 = back)
            chartOverlay.ToFrontOrBack(1);

            // Save the workbook
            workbook.Save("ChartOverlayToFront.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
