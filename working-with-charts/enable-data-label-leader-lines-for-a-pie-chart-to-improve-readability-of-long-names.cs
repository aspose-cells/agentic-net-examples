using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

class EnableLeaderLinesDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Long Category Name 1");
        sheet.Cells["A3"].PutValue("Long Category Name 2");
        sheet.Cells["A4"].PutValue("Long Category Name 3");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a pie chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure data labels
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;
        series.DataLabels.ShowPercentage = true;
        series.DataLabels.Position = LabelPositionType.OutsideEnd;

        // Enable leader lines for better readability of long category names
        series.HasLeaderLines = true;
        // Optional: customize the appearance of the leader lines
        series.LeaderLines.IsAuto = false;
        series.LeaderLines.Style = LineType.Solid;
        series.LeaderLines.WeightPt = 1.0;
        series.LeaderLines.Color = Color.DarkGray;

        // Save the workbook with the configured chart
        workbook.Save("PieChart_With_LeaderLines.xlsx");
    }
}