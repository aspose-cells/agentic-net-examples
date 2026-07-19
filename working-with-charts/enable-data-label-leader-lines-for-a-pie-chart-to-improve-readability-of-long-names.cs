// Title: Aspose.Cells for .NET: Add and style data‑label leader lines in a pie chart (C#)
// Description: Creates a workbook, inserts a pie chart with long category names, shows values and percentages, positions labels outside the slices, enables leader lines, and customizes their style, weight, and color before saving the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | pie chart | leader lines | data labels | outside label position | customize leader line style | Excel chart example | GitHub sample | Aspose.Cells chart tutorial
// Common Searches: Aspose.Cells enable leader lines pie chart | C# pie chart data label outside end Aspose.Cells | how to set leader line color and weight Aspose.Cells | customize leader line style solid Aspose.Cells | pie chart long category names leader lines .NET | Aspose.Cells chart examples GitHub
// Developer Intent: Add leader lines to pie‑chart data labels and adjust their visual properties using Aspose.Cells for .NET.
// Use Cases: Enhance readability of long category names by placing labels outside slices and linking them with leader lines. | Apply a uniform leader‑line style (solid, dark gray, 1 pt) across automated Excel reports. | Generate workbooks where pie charts automatically display both values and percentages with clearly styled connectors. | Create reusable chart templates for dashboards that require consistent label positioning and line formatting.
// AI Prompts: Write C# code with Aspose.Cells that builds a pie chart, shows values and percentages, positions labels outside, and adds dark‑gray solid leader lines of 1 pt. | Explain step‑by‑step how to enable and style leader lines for pie‑chart data labels in Aspose.Cells for .NET. | Provide a snippet that toggles the IsAuto property of leader lines and changes their style to solid for a given chart series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// Creates a workbook, inserts a pie chart with long category names, shows values and percentages, positions labels outside the slices, enables leader lines, and customizes their style, weight, and color before saving the file as an Excel workbook.
class EnableLeaderLines
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with long category names
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Long Category Name 1");
        worksheet.Cells["A3"].PutValue("Long Category Name 2");
        worksheet.Cells["A4"].PutValue("Long Category Name 3");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a pie chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 20, 12);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure data labels and enable leader lines
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;
        series.DataLabels.ShowPercentage = true;
        series.DataLabels.Position = LabelPositionType.OutsideEnd;
        series.HasLeaderLines = true; // Enable leader lines for better readability

        // Optional: customize the appearance of the leader lines
        series.LeaderLines.IsAuto = false;
        series.LeaderLines.Style = LineType.Solid;
        series.LeaderLines.WeightPt = 1.0;
        series.LeaderLines.Color = Color.DarkGray;

        // Save the workbook with the configured chart
        workbook.Save("PieChart_With_LeaderLines.xlsx");
    }
}
