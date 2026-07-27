using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartPositionDebug
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

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Calculate the chart so that shape positions are populated
            chart.Calculate();

            // Retrieve the first point of the first series
            ChartPoint point = chart.NSeries[0].Points[0];

            // Log chart object position (row/column based)
            Console.WriteLine("Chart UpperLeftRow: " + chart.ChartObject.UpperLeftRow);
            Console.WriteLine("Chart UpperLeftColumn: " + chart.ChartObject.UpperLeftColumn);
            Console.WriteLine("Chart Left (pixels): " + chart.ChartObject.Left);
            Console.WriteLine("Chart Top (pixels): " + chart.ChartObject.Top);

            // Log the point's position in pixels
            Console.WriteLine("Point ShapeXPx (X position in pixels): " + point.ShapeXPx);
            Console.WriteLine("Point ShapeYPx (Y position in pixels): " + point.ShapeYPx);

            // Also log the relative positions (1/4000 of chart width/height)
            Console.WriteLine("Point ShapeX (1/4000 of chart width): " + point.ShapeX);
            Console.WriteLine("Point ShapeY (1/4000 of chart height): " + point.ShapeY);

            // Save the workbook (debug workbook can be inspected if needed)
            workbook.Save("ChartPositionDebug.xlsx");
        }
    }
}