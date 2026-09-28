// Title: Change a column chart to an area chart while keeping existing data label settings in Aspose.Cells for .NET
// AI Prompts: Set Chart.Type from ChartType.Column to ChartType.Area and retain all DataLabels properties such as ShowValue, ShowCategoryName, ShapeType, and custom formatting. | Programmatically switch an Aspose.Cells chart to an area chart without resetting the configured data label colors, shape, and visibility.
// Common Searches: Aspose.Cells C# change chart type column to area preserve data labels | keep data label formatting when converting chart type in Aspose.Cells | how to retain custom data label shape after changing chart type Aspose.Cells | switch chart from column to area without losing label settings .NET
// Tags: Aspose.Cells chart type conversion | preserve data label settings Aspose.Cells | C# change chart to area chart | Aspose.Cells data label formatting | chart type change without resetting labels

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example creates a workbook, adds sample data, inserts a column chart, configures data labels (value, category name, rectangular shape, dark blue background, custom formatting), then changes the chart type to Area while preserving those label settings, and saves the file as ChartColumnToArea.xlsx.
class ChangeChartTypeExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart (initial type)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure data labels for the first series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;                     // Show values
        series.DataLabels.ShowCategoryName = true;              // Show category names
        series.DataLabels.ShapeType = DataLabelShapeType.Rect;  // Rectangular label shape
        series.DataLabels.Area.ForegroundColor = Color.DarkBlue;
        series.DataLabels.Area.Formatting = FormattingType.Custom;

        // Change the chart type to Area while preserving the data label settings
        chart.Type = ChartType.Area;

        // Save the workbook
        workbook.Save("ChartColumnToArea.xlsx", SaveFormat.Xlsx);
    }
}
