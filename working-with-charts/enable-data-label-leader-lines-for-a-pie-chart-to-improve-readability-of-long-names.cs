using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsLeaderLinesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Long Category Name A");
            sheet.Cells["A3"].PutValue("Long Category Name B");
            sheet.Cells["A4"].PutValue("Long Category Name C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series (the only one in this case)
            Series series = chart.NSeries[0];

            // Enable data labels and position them outside the pie slices
            series.DataLabels.ShowValue = true;
            series.DataLabels.ShowPercentage = true;
            series.DataLabels.Position = LabelPositionType.OutsideEnd;

            // Enable leader lines for the series
            series.HasLeaderLines = true;

            // Optional: customize the appearance of the leader lines
            series.LeaderLines.IsAuto = false;               // Use custom settings
            series.LeaderLines.Style = LineType.Dot;         // Dotted line style
            series.LeaderLines.WeightPt = 1.0;               // Line thickness in points
            series.LeaderLines.Color = Color.Blue;           // Line color

            // Save the workbook to a file
            workbook.Save("PieChart_With_LeaderLines.xlsx");
        }
    }
}