// Title: Add a Rectangle Shape Over a Chart Plot Area Using Aspose.Cells for .NET
// Description: This C# example creates a workbook, inserts sample data, builds a column chart, calculates the plot‑area dimensions, converts the ratios to the 1/4000 chart unit required by AddShapeInChart, places an external rectangle that aligns exactly with the plot area, applies simple formatting, and saves the file as an XLSX document.
// Keywords: Aspose.Cells | C# chart shape | AddShapeInChart | plot area coordinates | external shape | chart overlay | Aspose.Cells example | Excel automation .NET
// Common Searches: Aspose.Cells add shape to chart plot area | Get chart plot area position Aspose.Cells .NET | How to overlay rectangle on Excel chart using Aspose | Convert PlotArea ratios to AddShapeInChart units | Place external shape in Aspose.Cells chart
// Developer Intent: Insert an external rectangle that matches the chart's plot area and persist the workbook.
// Use Cases: Visually emphasize the data region of a chart in automated reports | Add branded or labeled annotations that line up with the plot area | Generate dynamic Excel files where shapes are positioned relative to chart dimensions
// AI Prompts: Write C# code with Aspose.Cells to place a circular shape over the plot area of a line chart and set its opacity. | Explain the formula for translating PlotArea.XRatioToChart, YRatioToChart, WidthRatioToChart, HeightRatioToChart into the integer values required by AddShapeInChart. | Provide a tutorial for adding multiple shapes to different chart sections (plot area, legend, title) using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExternalShapeExample
{
    // This C# example creates a workbook, inserts sample data, builds a column chart, calculates the plot‑area dimensions, converts the ratios to the 1/4000 chart unit required by AddShapeInChart, places an external rectangle that aligns exactly with the plot area, applies simple formatting, and saves the file as an XLSX document.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Ensure the chart is calculated so that PlotArea dimensions are valid
                chart.Calculate();

                // Get plot area ratios relative to the chart (values between 0 and 1)
                double leftRatio = chart.PlotArea.XRatioToChart;
                double topRatio = chart.PlotArea.YRatioToChart;
                double widthRatio = chart.PlotArea.WidthRatioToChart;
                double heightRatio = chart.PlotArea.HeightRatioToChart;

                // Convert ratios to the unit required by AddShapeInChart (1/4000 of chart area)
                int left = (int)(leftRatio * 4000);
                int top = (int)(topRatio * 4000);
                int right = (int)((leftRatio + widthRatio) * 4000);
                int bottom = (int)((topRatio + heightRatio) * 4000);

                // Add an external rectangle shape exactly over the plot area
                Shape shape = chart.Shapes.AddShapeInChart(
                    MsoDrawingType.Rectangle,   // Shape type
                    PlacementType.Move,        // Placement behavior
                    left, top, right, bottom);

                // Optional: set some visual properties
                shape.Fill.SolidFill.Color = System.Drawing.Color.LightBlue;
                shape.Line.SolidFill.Color = System.Drawing.Color.DarkBlue;
                shape.Text = "Plot Area Overlay";

                // Save the workbook
                string outputPath = "ChartWithExternalShape.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
