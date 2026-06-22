using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExamples
{
    public class StackedBarChartDataLabelsAndShapeResize
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
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");
            worksheet.Cells["A5"].PutValue("Q4");

            worksheet.Cells["B1"].PutValue("Product A");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["B3"].PutValue(40);
            worksheet.Cells["B4"].PutValue(20);
            worksheet.Cells["B5"].PutValue(10);

            worksheet.Cells["C1"].PutValue("Product B");
            worksheet.Cells["C2"].PutValue(20);
            worksheet.Cells["C3"].PutValue(30);
            worksheet.Cells["C4"].PutValue(25);
            worksheet.Cells["C5"].PutValue(15);

            // Add a stacked bar chart
            int chartIndex = worksheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Set data range for the chart
            chart.NSeries.Add("B2:C5", true);               // Values
            chart.NSeries.CategoryData = "A2:A5";           // Categories

            // Enable data labels to show both value and percentage for each series
            foreach (Series series in chart.NSeries)
            {
                series.DataLabels.ShowValue = true;          // Show cell value
                series.DataLabels.ShowPercentage = true;    // Show percentage

                // Prevent automatic resizing of the label shape so we can set a fixed size
                series.DataLabels.IsResizeShapeToFitText = false;

                // Set a fixed width for the data label shape (in pixels)
                series.DataLabels.WidthPixel = 80;
            }

            // Resize the chart shape itself (using points)
            ChartShape chartShape = chart.ChartObject;
            chartShape.Width = 800;   // Width in points
            chartShape.Height = 600;  // Height in points

            // Add a rectangle shape inside the chart area using scale coordinates
            Shape rectangle = chart.Shapes.AddShapeInChartByScale(
                MsoDrawingType.Rectangle,
                PlacementType.Move,
                0.10,   // left  = 10% from left edge of chart area
                0.10,   // top   = 10% from top edge of chart area
                0.30,   // right = 30% from left edge of chart area
                0.30);  // bottom= 30% from top edge of chart area

            // Customize the rectangle appearance
            rectangle.Fill.SolidFill.Color = Color.LightBlue;
            rectangle.Line.SolidFill.Color = Color.DarkBlue;
            rectangle.Line.Weight = 1.5;

            // Save the workbook
            string outputPath = "StackedBarChart_WithDataLabelsAndResizedShapes.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}