// Title: Place a Shape Over a Chart Plot Area Using Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add sample data, generate a column chart, calculate its layout, obtain chart and plot‑area dimensions, compute relative percentages, insert a rectangle shape that exactly matches the plot area with AddShapeInChartByScale, style the shape, and save the file.
// Keywords: Aspose.Cells C# chart shape | AddShapeInChartByScale example | chart plot area dimensions | overlay shape on chart | save workbook Aspose.Cells | rectangle shape chart Aspose | chart area scaling | Aspose.Cells drawing API | place shape by percent | C# Excel chart annotation
// Common Searches: Aspose.Cells add rectangle to chart plot area | How to use AddShapeInChartByScale in .NET | Get plot area size from Aspose.Cells chart | Overlay shape on Excel chart with Aspose.Cells | C# code to align shape with chart plot area | Save workbook after adding chart shape Aspose | Aspose.Cells chart drawing examples
// Developer Intent: Insert a shape that aligns precisely with a chart’s plot area and persist the workbook.
// Use Cases: Emphasize data region with a semi‑transparent overlay in financial reports. | Add a branded watermark or logo confined to the plot area. | Create an interactive hotspot over the plot area for drill‑down actions. | Mask unwanted chart elements by covering the plot area. | Generate a custom grid or background within the plot area.
// AI Prompts: Write C# code with Aspose.Cells that adds a rectangle shape covering the plot area of a column chart and saves the workbook. | Show how to calculate plot‑area width and height as percentages of the chart and use AddShapeInChartByScale to position a shape. | Provide an Aspose.Cells example that styles the shape with semi‑transparent fill and dark border. | Explain how to retrieve chart dimensions and overlay an image on the plot area in C#. | Generate a step‑by‑step guide for placing any drawing object inside a chart plot area using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add sample data, generate a column chart, calculate its layout, obtain chart and plot‑area dimensions, compute relative percentages, insert a rectangle shape that exactly matches the plot area with AddShapeInChartByScale, style the shape, and save the file.
    public class PlaceShapeAtPlotArea
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(1500);
            worksheet.Cells["B3"].PutValue(2300);
            worksheet.Cells["B4"].PutValue(1800);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];
            chart.SetChartDataRange("A1:B4", true);
            chart.Calculate(); // Ensure the chart layout is calculated

            // Retrieve chart dimensions (in points) via ChartObject
            double chartWidth = chart.ChartObject.Width;
            double chartHeight = chart.ChartObject.Height;

            // Retrieve plot area dimensions (in points)
            double plotWidth = chart.PlotArea.Width;
            double plotHeight = chart.PlotArea.Height;

            // Plot area starts at the top‑left corner of the chart area (percent = 0)
            double leftPercent = 0.0;
            double topPercent = 0.0;
            double rightPercent = plotWidth / chartWidth;   // width as fraction of chart width
            double bottomPercent = plotHeight / chartHeight; // height as fraction of chart height

            // Add a rectangle shape that exactly covers the plot area
            Shape shape = chart.Shapes.AddShapeInChartByScale(
                MsoDrawingType.Rectangle,   // shape type
                PlacementType.Move,         // placement behavior
                leftPercent,                // left (percent of chart width)
                topPercent,                 // top (percent of chart height)
                rightPercent,               // right (percent of chart width)
                bottomPercent);             // bottom (percent of chart height)

            // Optional: format the shape to make it visible
            shape.Fill.SolidFill.Color = System.Drawing.Color.FromArgb(128, System.Drawing.Color.LightGreen); // semi‑transparent fill
            shape.Line.SolidFill.Color = System.Drawing.Color.DarkGreen; // border color

            // Save the workbook
            string outputPath = "ChartWithPlotAreaShape.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
