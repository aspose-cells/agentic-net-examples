using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartShapeConnectionPoints
{
    class Program
    {
        static void Main()
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
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add a rectangle shape inside the chart (this shape will have connection points)
            Shape shape = chart.Shapes.AddShape(MsoDrawingType.Rectangle, 1000, 1000, 2000, 1000, 0, 0);
            shape.Text = "Sample Shape";

            // Retrieve connection points of the shape
            float[][] points = shape.GetConnectionPoints();

            // Export the connection points to a CSV file
            string csvPath = "ChartShapeConnectionPoints.csv";
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Write header
                writer.WriteLine("PointIndex,X,Y");
                // Write each point
                for (int i = 0; i < points.Length; i++)
                {
                    float x = points[i][0];
                    float y = points[i][1];
                    writer.WriteLine($"{i + 1},{x},{y}");
                }
            }

            // Save the workbook (optional, to visualize the chart and shape)
            workbook.Save("ChartShapeConnectionPoints.xlsx");
        }
    }
}