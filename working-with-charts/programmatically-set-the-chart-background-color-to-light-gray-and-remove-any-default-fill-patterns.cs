using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartBackground
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(20);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ----- Set chart background to light gray -----
            // Chart area background
            chart.ChartArea.Area.BackgroundColor = Color.LightGray;
            chart.ChartArea.BackgroundMode = BackgroundMode.Opaque; // ensure the color is shown
            // Remove any fill pattern from the chart area
            chart.ChartArea.Area.FillFormat.Pattern = FillPattern.None;

            // Plot area background (optional, often visible inside the chart)
            chart.PlotArea.Area.BackgroundColor = Color.LightGray;
            chart.PlotArea.BackgroundMode = BackgroundMode.Opaque;
            chart.PlotArea.Area.FillFormat.Pattern = FillPattern.None;

            // Save the workbook
            workbook.Save("ChartBackgroundLightGray.xlsx");
        }
    }
}