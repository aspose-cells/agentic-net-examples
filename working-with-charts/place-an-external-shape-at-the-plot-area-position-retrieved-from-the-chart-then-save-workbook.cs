using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartShapeExample
{
    class Program
    {
        static void Main()
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

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Ensure the chart is calculated so that PlotArea dimensions are valid
                chart.Calculate();

                // Convert plot area size to the unit required by AddShapeInChart (1/4000 of chart area)
                // Plot area starts at the left/top edge of the chart (0,0)
                int left = 0;
                int top = 0;
                int right = (int)(chart.PlotArea.WidthRatioToChart * 4000);
                int bottom = (int)(chart.PlotArea.HeightRatioToChart * 4000);

                // Add a rectangle shape that covers the plot area
                Shape shape = chart.Shapes.AddShapeInChart(
                    MsoDrawingType.Rectangle,
                    PlacementType.Move,
                    left,
                    top,
                    right,
                    bottom);

                // Optional: set some text for the shape
                shape.Text = "Plot Area";

                // Save the workbook
                workbook.Save("ChartWithPlotAreaShape.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}