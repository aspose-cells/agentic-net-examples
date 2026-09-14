// Title: Enable and customize data label leader lines for a pie chart using Aspose.Cells in C#
// AI Prompts: Generate C# code with Aspose.Cells that creates a pie chart, shows values and percentages outside the slices, and activates leader lines. | Update an existing Aspose.Cells pie chart to programmatically set leader line style, thickness, and color.
// Common Searches: Aspose.Cells C# pie chart leader lines outside data labels | how to set custom leader line color and weight in Aspose.Cells pie chart | display percentage and value labels with leader lines in Aspose.Cells .NET | C# example for enabling leader lines on pie chart series using Aspose.Cells
// Tags: pie chart leader lines Aspose.Cells C# | outside data label position Aspose.Cells | customize leader line style Aspose.Cells | set leader line thickness Aspose.Cells | display percentages in pie chart labels Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsLeaderLinesDemo
{
    // Creates a workbook, adds sample data, inserts a pie chart, configures the series to show values and percentages outside the slices, enables leader lines, customizes their style, weight, and color, and saves the file as PieChart_With_LeaderLines.xlsx.
    class Program
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
            sheet.Cells["B2"].PutValue(40);
            sheet.Cells["B3"].PutValue(35);
            sheet.Cells["B4"].PutValue(25);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first (and only) series
            Series series = chart.NSeries[0];

            // Enable data labels and position them outside the pie slices
            series.DataLabels.ShowValue = true;
            series.DataLabels.ShowPercentage = true;
            series.DataLabels.Position = LabelPositionType.OutsideEnd;

            // Enable leader lines for the series
            series.HasLeaderLines = true;

            // Optional: customize the appearance of the leader lines
            series.LeaderLines.IsAuto = false;               // Use custom settings
            series.LeaderLines.Style = LineType.Solid;       // Solid line style
            series.LeaderLines.WeightPt = 1.0;               // Line thickness in points
            series.LeaderLines.Color = Color.DarkGray;       // Line color

            // Save the workbook to a file
            workbook.Save("PieChart_With_LeaderLines.xlsx");
        }
    }
}
