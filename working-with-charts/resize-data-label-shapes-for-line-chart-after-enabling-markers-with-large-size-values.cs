// Title: Resize Data Label Shapes in a Line Chart with Large Markers – Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to create a line chart, enable 30‑point circle markers, turn off auto‑fit, and set each data label’s width to 80 pt and height to 30 pt using Aspose.Cells for .NET, then recalculate and save the workbook.
// Keywords: Aspose.Cells | C# | line chart | data label size | custom data label dimensions | large marker size | disable auto fit | ChartPoint.DataLabels | ResizeShapeToFitText | ChartMarkerType.Circle
// Common Searches: Aspose.Cells resize data label shape | Set custom width and height for chart data labels .NET | How to disable auto‑fit for data labels in Aspose.Cells | Change data label size after increasing marker size | C# line chart data label dimensions Aspose.Cells
// Developer Intent: Set explicit width and height for each data label after enabling large markers in a line chart.
// Use Cases: Generate a line chart with X/Y series and 30‑point circle markers. | Show value labels centered on each point and turn off automatic shape fitting. | Assign a fixed width of 80 pt and height of 30 pt to every data label shape. | Recalculate the chart to apply changes and save the workbook.
// AI Prompts: Write C# code using Aspose.Cells to create a line chart with 30‑point circle markers and resize each data label to 80 pt × 30 pt. | Explain how to disable the auto‑fit feature for chart data labels and set custom dimensions for individual points in Aspose.Cells for .NET. | Demonstrate adjusting data label positions after resizing shapes in a line chart with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to create a line chart, enable 30‑point circle markers, turn off auto‑fit, and set each data label’s width to 80 pt and height to 30 pt using Aspose.Cells for .NET, then recalculate and save the workbook.
    class ResizeDataLabelShapes
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the line chart
                worksheet.Cells["A1"].PutValue("X");
                worksheet.Cells["B1"].PutValue("Y");
                for (int i = 0; i < 10; i++)
                {
                    worksheet.Cells[i + 2, 0].PutValue(i + 1);          // X values
                    worksheet.Cells[i + 2, 1].PutValue((i + 1) * 2);   // Y values
                }

                // Add a line chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Add the data series to the chart
                int seriesIndex = chart.NSeries.Add("B2:B11", true);
                Series series = chart.NSeries[seriesIndex];
                series.XValues = "A2:A11";

                // Enable markers and set a large marker size
                series.Marker.MarkerStyle = ChartMarkerType.Circle;
                series.Marker.MarkerSize = 30; // size in points (large)

                // Show data labels for the series
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;

                // Resize each data label shape
                foreach (ChartPoint point in series.Points)
                {
                    // Disable auto‑fit so custom dimensions are applied
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Set custom width and height (in points)
                    point.DataLabels.Width = 80;
                    point.DataLabels.Height = 30;
                }

                // Recalculate the chart to apply the changes
                chart.Calculate();

                // Save the workbook
                string outputPath = "LineChartDataLabelResize.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ResizeDataLabelShapes.Run();
        }
    }
}
