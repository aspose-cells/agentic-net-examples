using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartPointColorDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Iterate over each series in the chart
            foreach (Series series in chart.NSeries)
            {
                // Iterate over each point in the current series
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];

                    // Assign a custom foreground color to the point's area.
                    // Example: even-indexed points get Red, odd-indexed points get Blue.
                    point.Area.ForegroundColor = (i % 2 == 0) ? Color.Red : Color.Blue;

                    // Ensure the custom formatting is applied
                    point.Area.Formatting = FormattingType.Custom;
                }
            }

            // Save the workbook
            workbook.Save("ChartPointCustomColors.xlsx");
        }
    }
}