// Title: Add Leader Lines to Pie Chart Data Labels with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts sample data with long category names, adds a pie chart, shows values and category names, positions labels outside the slices, enables leader lines, and customizes their style, weight, and color before saving the file.
// Keywords: Aspose.Cells pie chart leader lines | C# data label leader lines | customize chart label lines Aspose | outside label position pie chart | Excel export pie chart formatting
// Common Searches: how to enable leader lines for pie chart labels Aspose.Cells | set leader line color and thickness in Aspose.Cells chart | position pie chart data labels outside with lines | Aspose.Cells C# pie chart label formatting examples
// Developer Intent: Enable and style leader lines for pie‑chart data labels using Aspose.Cells in a .NET application.
// Use Cases: Clarify long category names on a pie chart by placing labels outside and connecting them with leader lines. | Produce presentation‑ready Excel reports where pie‑chart labels are visually distinct and easy to read. | Programmatically control label line appearance (style, weight, color) for consistent branding across generated workbooks.
// AI Prompts: Write C# code with Aspose.Cells that adds a pie chart, shows category names and values, places labels outside, and turns on leader lines. | Show how to toggle leader lines for a chart series based on a boolean variable in Aspose.Cells. | Explain how to batch‑update leader line thickness and style for all series in an Aspose.Cells workbook.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, inserts sample data with long category names, adds a pie chart, shows values and category names, positions labels outside the slices, enables leader lines, and customizes their style, weight, and color before saving the file.
class EnableLeaderLines
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with long category names
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Very Long Category Name 1");
        worksheet.Cells["A3"].PutValue("Very Long Category Name 2");
        worksheet.Cells["A4"].PutValue("Very Long Category Name 3");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a pie chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the first series and configure data labels
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;                     // Show numeric values
        series.DataLabels.ShowCategoryName = true;              // Show category names
        series.DataLabels.Position = LabelPositionType.OutsideEnd; // Position labels outside

        // Enable leader lines to connect labels with their slices
        series.HasLeaderLines = true;

        // Optional: customize the appearance of the leader lines
        series.LeaderLines.IsAuto = false;                      // Disable automatic formatting
        series.LeaderLines.Style = LineType.Solid;              // Solid line style
        series.LeaderLines.WeightPt = 1.0;                      // Line thickness in points
        series.LeaderLines.Color = Color.DarkGray;              // Line color

        // Save the workbook with the configured chart
        workbook.Save("PieChartWithLeaderLines.xlsx");
    }
}
