// Title: Rotate Primary Y‑Axis Tick Labels 90° Clockwise with Aspose.Cells (C#)
// Description: Demonstrates how to access the primary Y‑axis (value axis) of a chart in Aspose.Cells, set its TickLabels.RotationAngle to 90 degrees, and save the workbook. The example creates sample data, builds a column chart, rotates the Y‑axis labels vertically, and outputs an XLSX file.
// Keywords: Aspose.Cells rotate Y axis labels | C# chart tick label rotation | TickLabels.RotationAngle Aspose | primary value axis label orientation | Aspose.Cells chart formatting | .NET chart axis label angle | vertical Y‑axis tick labels | Aspose.Cells example rotate labels
// Common Searches: rotate Y axis tick labels Aspose.Cells C# | set chart axis label angle Aspose.Cells .NET | how to change tick label orientation in Aspose chart | Aspose.Cells primary value axis label rotation | C# make Y‑axis labels vertical in Excel chart
// Developer Intent: Apply a 90‑degree clockwise rotation to the primary Y‑axis tick labels of an Aspose.Cells chart.
// Use Cases: Enhance readability when numeric tick marks are crowded on a column chart. | Create a consistent vertical label style for dashboards generated programmatically. | Prepare Excel reports where Y‑axis labels must align with design guidelines.
// AI Prompts: Show C# code that rotates the primary Y‑axis tick labels 90 degrees clockwise using Aspose.Cells. | Explain how to access and modify TickLabels properties for a chart's value axis in Aspose.Cells .NET. | Provide a step‑by‑step guide to set TickLabels.RotationAngle for both primary and secondary axes in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTickLabelsRotation
{
    // Demonstrates how to access the primary Y‑axis (value axis) of a chart in Aspose.Cells, set its TickLabels.RotationAngle to 90 degrees, and save the workbook. The example creates sample data, builds a column chart, rotates the Y‑axis labels vertically, and outputs an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the primary Y axis (value axis) tick labels
            TickLabels yAxisTickLabels = chart.ValueAxis.TickLabels;

            // Rotate the tick labels 90 degrees clockwise
            yAxisTickLabels.RotationAngle = 90;

            // Save the workbook
            workbook.Save("PrimaryYAxisTickLabelsRotation.xlsx");
        }
    }
}
