// Title: How to format the second series data labels in an Aspose.Cells column chart with bold red font and yellow background using C#
// AI Prompts: Generate C# code that accesses the second series of an Aspose.Cells column chart and applies bold red text with a yellow background to its data labels. | Show how to programmatically set custom font styling and background fill for data labels of a specific series in an Aspose.Cells chart.
// Common Searches: C# Aspose.Cells set bold red font for data labels of second series in column chart | Aspose.Cells change background fill of chart data labels for a particular series | How to customize font color and label background in an Aspose.Cells column chart using .NET
// Tags: chart data label text style Aspose.Cells C# | chart data label background shading Aspose.Cells C# | column chart series label customization .NET | Aspose.Cells data label appearance customization

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds two data series, builds a column chart, enables data labels for the second series, and then formats those labels with a bold red font and a yellow background before saving the file.
class FormatSecondSeriesDataLabels
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for two series
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["A5"].PutValue("D");

        // First series values
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(40);

        // Second series values
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);
        sheet.Cells["C5"].PutValue(45);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add both series to the chart
        chart.NSeries.Add("B2:B5", true); // first series
        chart.NSeries.Add("C2:C5", true); // second series
        chart.NSeries.CategoryData = "A2:A5";

        // Enable data labels for the second series
        Series secondSeries = chart.NSeries[1];
        secondSeries.DataLabels.ShowValue = true;

        // Apply bold font and red color
        secondSeries.DataLabels.Font.IsBold = true;
        secondSeries.DataLabels.Font.Color = Color.Red;

        // Set yellow background for the data labels
        secondSeries.DataLabels.Area.BackgroundColor = Color.Yellow;
        secondSeries.DataLabels.BackgroundMode = BackgroundMode.Automatic; // ensure background is visible

        // Apply the font settings to all data label nodes
        secondSeries.DataLabels.ApplyFont();

        // Save the workbook
        workbook.Save("FormattedSecondSeriesDataLabels.xlsx");
    }
}
