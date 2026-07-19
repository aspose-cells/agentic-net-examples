// Title: C# – Rotate Primary Y‑Axis Tick Labels 90° Clockwise in Aspose.Cells Chart
// Description: This example creates a workbook, adds sample data, inserts a column chart, disables automatic rotation on the primary Y‑axis (value axis) tick labels, sets a manual RotationAngle of 90 degrees clockwise, and saves the file as YAxisTickLabelsRotated.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells rotate Y axis labels | C# chart tick label rotation | ValueAxis TickLabels RotationAngle | disable automatic tick label rotation | Aspose.Cells .NET chart formatting
// Common Searches: rotate Y‑axis tick labels 90 degrees Aspose.Cells | Aspose.Cells C# set chart axis label angle | how to change tick label rotation in Aspose.Cells chart | manual rotation of chart axis labels .NET
// Developer Intent: Apply a 90° clockwise rotation to the primary Y‑axis tick labels in an Aspose.Cells chart.
// Use Cases: Improve readability of long category names on the Y‑axis of column charts. | Create compact reports where vertical space is limited. | Override default automatic rotation when custom label orientation is required.
// AI Prompts: Show C# code to set a 90° clockwise rotation for primary Y‑axis tick labels in Aspose.Cells. | Give an Aspose.Cells .NET example that disables automatic rotation and applies a custom RotationAngle to chart axis TickLabels. | Explain how to adjust TickLabels.RotationAngle for both primary and secondary axes in an Aspose.Cells chart using C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a workbook, adds sample data, inserts a column chart, disables automatic rotation on the primary Y‑axis (value axis) tick labels, sets a manual RotationAngle of 90 degrees clockwise, and saves the file as YAxisTickLabelsRotated.xlsx using Aspose.Cells for .NET.
class RotateYaxisTickLabels
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category 1");
        sheet.Cells["A2"].PutValue("Category 2");
        sheet.Cells["A3"].PutValue("Category 3");
        sheet.Cells["B1"].PutValue(100);
        sheet.Cells["B2"].PutValue(200);
        sheet.Cells["B3"].PutValue(300);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data source
        chart.NSeries.Add("B1:B3", true);
        chart.NSeries.CategoryData = "A1:A3";

        // Access the primary Y axis (value axis) tick labels
        // Disable automatic rotation and set a manual rotation of 90 degrees clockwise
        chart.ValueAxis.TickLabels.IsAutomaticRotation = false;
        chart.ValueAxis.TickLabels.RotationAngle = 90;

        // Save the workbook
        workbook.Save("YAxisTickLabelsRotated.xlsx");
    }
}
