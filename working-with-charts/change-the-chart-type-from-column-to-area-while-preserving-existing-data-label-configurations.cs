// Title: Change a Column Chart to an Area Chart while retaining data label settings – Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds a column chart with sample data, configures data labels (value, category name, rectangular shape, custom area formatting, light‑green fill), then switches the chart type to Area using chart.Type. The data‑label properties remain unchanged, and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | chart type conversion | Column to Area | preserve data labels | ChartType.Area | data label formatting | Excel chart automation | Aspose.Cells for .NET
// Common Searches: Aspose.Cells change chart type column to area | keep data label formatting when changing chart type Aspose.Cells | C# Aspose.Cells preserve chart data labels | convert column chart to area chart programmatically | Aspose.Cells ChartType.Area example
// Developer Intent: Switch an existing column chart to an area chart without resetting its data‑label configuration.
// Use Cases: Generate a report where the chart style can be toggled between column and area while maintaining custom label appearance. | Provide end‑users the ability to select a different chart type at runtime without re‑applying label settings. | Reuse a chart template with predefined data‑label formatting and programmatically change its type for varied datasets.
// AI Prompts: Write C# code using Aspose.Cells to change a chart from Column to Area while keeping all data label properties. | Show how to modify ChartType in an Aspose.Cells workbook without losing data label shape, color, or visibility. | Explain step‑by‑step how to switch a chart to Area in Aspose.Cells and retain existing data label configurations.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, adds a column chart with sample data, configures data labels (value, category name, rectangular shape, custom area formatting, light‑green fill), then switches the chart type to Area using chart.Type. The data‑label properties remain unchanged, and the workbook is saved as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
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

        // Add a column chart (initial type) to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure data labels for the first series (these settings should be preserved)
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;
        series.DataLabels.ShowCategoryName = true;
        series.DataLabels.ShapeType = DataLabelShapeType.Rect;
        series.DataLabels.Area.Formatting = FormattingType.Custom;
        series.DataLabels.Area.ForegroundColor = System.Drawing.Color.LightGreen;

        // Change the chart type from Column to Area while keeping data label settings intact
        chart.Type = ChartType.Area;

        // Save the workbook to a file
        workbook.Save("ChartColumnToArea.xlsx", SaveFormat.Xlsx);
    }
}
