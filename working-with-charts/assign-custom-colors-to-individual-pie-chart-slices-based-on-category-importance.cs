using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsPieCustomColors
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: categories and their importance values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("High");
            sheet.Cells["A3"].PutValue("Medium");
            sheet.Cells["A4"].PutValue("Low");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);   // High importance
            sheet.Cells["B3"].PutValue(30);   // Medium importance
            sheet.Cells["B4"].PutValue(20);   // Low importance

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Define custom colors based on importance
            // High -> Red, Medium -> Orange, Low -> Green
            Color[] sliceColors = new Color[] { Color.Red, Color.Orange, Color.Green };

            // Apply custom colors to each slice (chart point)
            Series series = chart.NSeries[0];
            for (int i = 0; i < series.Points.Count; i++)
            {
                ChartPoint point = series.Points[i];
                point.Area.ForegroundColor = sliceColors[i];
                point.Area.Formatting = FormattingType.Custom; // Ensure custom color is used
            }

            // Save the workbook
            workbook.Save("PieChartCustomColors.xlsx");
        }
    }
}