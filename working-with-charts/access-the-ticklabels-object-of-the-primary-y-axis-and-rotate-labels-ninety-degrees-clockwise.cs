// Title: Aspose.Cells C# – Rotate Primary Y‑Axis Tick Labels 90° Clockwise
// Description: A concise C# sample that builds a workbook, inserts a column chart, and applies a 90‑degree clockwise rotation to the primary Y‑axis (value axis) tick labels via the ValueAxis.TickLabels.RotationAngle property before saving the file.
// Keywords: Aspose.Cells rotate Y axis labels | C# chart tick label angle | ValueAxis.TickLabels rotation | set chart axis label orientation .NET | vertical Y‑axis labels Aspose.Cells
// Common Searches: how to rotate Y‑axis tick labels 90 degrees Aspose.Cells | Aspose.Cells C# chart label rotation example | ValueAxis.TickLabels.RotationAngle usage | rotate primary value axis labels clockwise | Aspose.Cells chart label orientation .NET
// Developer Intent: Apply a 90° clockwise rotation to the primary Y‑axis tick labels in an Aspose.Cells chart using C#.
// Use Cases: Improve readability of long numeric Y‑axis labels by displaying them vertically. | Fit Y‑axis labels into narrow chart areas for compact PDF or printed reports. | Standardize vertical label orientation across multiple charts in a workbook.
// AI Prompts: Generate C# code with Aspose.Cells that rotates the primary Y‑axis tick labels 90° clockwise for any chart type. | Show how to set independent rotation angles for X‑axis and Y‑axis tick labels in an Aspose.Cells workbook. | Provide an example that rotates Y‑axis tick labels and also adjusts their font size to avoid overlap.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// A concise C# sample that builds a workbook, inserts a column chart, and applies a 90‑degree clockwise rotation to the primary Y‑axis (value axis) tick labels via the ValueAxis.TickLabels.RotationAngle property before saving the file.
class RotateYTickLabels
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
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

        // Access the primary Y axis (ValueAxis) tick labels and rotate them 90 degrees clockwise
        chart.ValueAxis.TickLabels.RotationAngle = 90;

        // Save the workbook to a file
        workbook.Save("YTickLabelsRotated.xlsx");
    }
}
