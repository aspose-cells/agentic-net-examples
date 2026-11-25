// Title: Add leader lines to pie chart data labels for long category names using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to show leader lines on pie chart labels | display long category names on pie chart with leader lines in .NET | customize pie chart data label leader line style using Aspose.Cells | enable outside label position with leader lines for pie chart in C# | set leader line color and thickness for Aspose.Cells pie chart
// Tags: pie chart leader lines Aspose.Cells | data label outside end Aspose.Cells | leader line appearance Aspose.Cells | long category names chart labels .NET | pie chart data label percentage Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a pie chart with long‑text categories, configures data labels to show category names and percentages, enables leader lines, customizes their style (solid, 1 pt, dark gray), and saves the file as PieChartWithLeaderLines.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data with long category names to demonstrate leader lines
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");

        sheet.Cells["A2"].PutValue("Very Long Category Name One");
        sheet.Cells["A3"].PutValue("Extremely Long Category Name Two");
        sheet.Cells["A4"].PutValue("Another Very Long Category Name Three");

        sheet.Cells["B2"].PutValue(40);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(30);

        // Add a pie chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure data labels to show category names and percentages
        Series series = chart.NSeries[0];
        series.DataLabels.ShowCategoryName = true;
        series.DataLabels.ShowPercentage = true;
        series.DataLabels.Position = LabelPositionType.OutsideEnd;

        // Enable leader lines so long category names are readable
        series.HasLeaderLines = true;

        // Optional: customize the appearance of the leader lines
        series.LeaderLines.IsAuto = false;               // Turn off automatic formatting
        series.LeaderLines.Style = LineType.Solid;       // Solid line
        series.LeaderLines.WeightPt = 1.0;               // 1 point thickness
        series.LeaderLines.Color = Color.DarkGray;       // Dark gray color

        // Save the workbook to a file
        string outputPath = "PieChartWithLeaderLines.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
    }
}
