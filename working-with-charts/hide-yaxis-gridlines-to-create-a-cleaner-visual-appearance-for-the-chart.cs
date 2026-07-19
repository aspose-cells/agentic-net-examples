// Title: Hide Y‑Axis Gridlines in a Column Chart with Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to create a workbook, add sample data, insert a column chart, and suppress the Y‑axis (value axis) major gridlines by setting chart.ValueAxis.MajorGridLines.IsVisible to false. The resulting Excel file shows a cleaner visual layout suitable for reports and dashboards.
// Keywords: Aspose.Cells C# | .NET chart formatting | hide Y axis gridlines | disable major gridlines | ValueAxis gridlines visibility | Excel column chart styling | Aspose.Cells chart API | remove chart gridlines
// Common Searches: Aspose.Cells hide Y axis gridlines C# | disable major gridlines on value axis Aspose.Cells | C# remove Y‑axis lines from column chart | chart.ValueAxis.MajorGridLines.IsVisible false example | how to hide Excel chart gridlines using Aspose.Cells
// Developer Intent: The developer needs to turn off the Y‑axis major gridlines of a column chart to achieve a minimalist appearance.
// Use Cases: Generating sales reports where charts are free of background gridlines for a sleek look. | Building financial dashboards that emphasize data trends without axis clutter. | Exporting analytics to Excel while complying with corporate style guides that forbid gridlines.
// AI Prompts: Provide C# code to hide both major and minor Y‑axis gridlines in an Aspose.Cells chart. | Show how to toggle Y‑axis gridline visibility with a boolean parameter in Aspose.Cells for .NET. | Explain how to customize the value axis line color and thickness after disabling its gridlines.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example demonstrates how to create a workbook, add sample data, insert a column chart, and suppress the Y‑axis (value axis) major gridlines by setting chart.ValueAxis.MajorGridLines.IsVisible to false. The resulting Excel file shows a cleaner visual layout suitable for reports and dashboards.
class HideYAxisGridlines
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Insert a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the chart data range
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Hide Y‑axis (value axis) major gridlines for a cleaner visual appearance
        chart.ValueAxis.MajorGridLines.IsVisible = false;

        // Save the workbook
        workbook.Save("HideYAxisGridlines.xlsx");
    }
}
