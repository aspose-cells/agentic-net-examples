// Title: Check if leader lines are enabled on a pie chart after moving data labels with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a pie chart, sets data labels to OutsideEnd, shifts them 10 pixels horizontally, enables leader lines, and prints the series.HasLeaderLines flag. | Write a method using Aspose.Cells to offset chart data labels, turn on leader lines for a series, and return a boolean indicating whether leader lines are visible. | Provide a step‑by‑step C# example that confirms the HasLeaderLines property after adjusting label positions in a pie chart.
// Common Searches: Aspose.Cells how to verify leader lines after changing data label position in C# | C# code to enable and check leader lines on a pie chart using Aspose.Cells | Does moving data labels affect HasLeaderLines property in Aspose.Cells charts? | Sample Aspose.Cells .NET example for leader line visibility on a pie chart series
// Tags: Aspose.Cells enable leader lines pie chart | C# Aspose.Cells offset data labels | check HasLeaderLines property Aspose.Cells | pie chart data label positioning Aspose.Cells | Aspose.Cells chart series leader lines

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a pie chart with sample data, configures data labels to show values outside the slices, offsets the labels 10 pixels on the X axis, enables leader lines for the series, reads the HasLeaderLines property to confirm they are enabled, prints the result, and saves the workbook as LeaderLinesVerification.xlsx.
class VerifyLeaderLines
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
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

        // Access the first series
        Series series = chart.NSeries[0];

        // Enable data labels and offset them (e.g., move 10 pixels to the right)
        series.DataLabels.ShowValue = true;
        series.DataLabels.Position = LabelPositionType.OutsideEnd;
        series.DataLabels.XPixel += 10; // offset

        // Enable leader lines for the series
        series.HasLeaderLines = true;

        // Verify leader lines visibility
        bool leaderLinesVisible = series.HasLeaderLines;
        Console.WriteLine("Leader lines enabled: " + leaderLinesVisible);

        // Save the workbook
        workbook.Save("LeaderLinesVerification.xlsx");
    }
}
