// Title: Set radar chart data label font size to 12 pt and color to blue using Aspose.Cells for .NET (C#)
// AI Prompts: Create a radar chart in C# with Aspose.Cells and format its data labels to 12‑point blue text. | Update an existing Aspose.Cells radar chart so that every series data label uses a 12 pt font and blue color. | Programmatically apply font size and color styling to the data labels of a radar chart series with Aspose.Cells for .NET.
// Common Searches: how to change data label font size and color in a radar chart using Aspose.Cells C# | Aspose.Cells C# set radar chart data labels to blue 12 point font | formatting data labels on radar charts with Aspose.Cells .NET example
// Tags: Aspose.Cells radar chart data label styling | C# set data label font size Aspose.Cells | Aspose.Cells chart series label color blue | modify radar chart label appearance .NET | apply font to chart data labels Aspose

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates creating a radar chart, adding a series, enabling data labels, and setting their font size to 12 points and color to blue with Aspose.Cells for .NET (C#).
class RadarChartDataLabelDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the radar chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Cat1");
        sheet.Cells["A3"].PutValue("Cat2");
        sheet.Cells["A4"].PutValue("Cat3");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(4);
        sheet.Cells["B3"].PutValue(2);
        sheet.Cells["B4"].PutValue(5);

        // Add a radar chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        int seriesIndex = chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the series
        Series series = chart.NSeries[seriesIndex];
        series.DataLabels.ShowValue = true;

        // Change data label font size to 12 points and color to blue
        series.DataLabels.Font.Size = 12;
        series.DataLabels.Font.Color = Color.Blue;

        // Apply the font settings to all data labels
        series.DataLabels.ApplyFont();

        // Save the workbook
        workbook.Save("RadarChartDataLabels.xlsx");
    }
}
