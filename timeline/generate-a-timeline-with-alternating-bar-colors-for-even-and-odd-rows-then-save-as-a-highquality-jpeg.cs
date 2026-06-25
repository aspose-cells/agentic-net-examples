using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsTimelineBarChart
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (Category and Value)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            string[] categories = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
            int[] values = { 120, 150, 90, 200, 130, 170 };

            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 1, 0].PutValue(categories[i]);   // Column A
                cells[i + 1, 1].PutValue(values[i]);      // Column B
            }

            // Add a column (bar) chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set data source for the chart
            chart.NSeries.Add("B2:B7", true);          // Values
            chart.NSeries.CategoryData = "A2:A7";      // Categories

            // Apply alternating colors to each data point
            // Even index (0‑based) -> LightBlue, Odd index -> LightCoral
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];
                // Ensure the point has a fill format
                point.Area.ForegroundColor = (i % 2 == 0) ? Color.LightBlue : Color.LightCoral;
            }

            // Save the chart as a high‑quality JPEG (quality = 100)
            string outputPath = "TimelineBarChart.jpg";
            chart.ToImage(outputPath, 100L);   // JPEG quality parameter

            // Optionally, save the workbook for reference
            workbook.Save("TimelineBarChartWorkbook.xlsx");
        }
    }
}