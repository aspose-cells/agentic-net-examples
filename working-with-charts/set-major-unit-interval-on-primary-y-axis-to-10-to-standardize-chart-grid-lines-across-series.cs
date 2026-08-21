// Title: Aspose.Cells for .NET: Set Primary Y‑Axis Major Unit to 10 in a Column Chart
// Description: Shows how to create a workbook, populate sample data, add a column chart, turn off automatic major unit, set the primary Y‑axis (value axis) major unit to 10, and save the result as ChartWithMajorUnit.xlsx using Aspose.Cells for .NET (C#).
// Keywords: Aspose.Cells | C# | .NET | chart major unit | Y axis interval | set major unit | disable automatic major unit | column chart formatting | Excel chart grid lines | ValueAxis.MajorUnit
// Common Searches: Aspose.Cells set Y axis major unit | C# chart major unit 10 Aspose | disable automatic major unit Aspose.Cells | standardize chart grid lines Aspose | how to set ValueAxis.MajorUnit in .NET
// Developer Intent: Configure the primary Y‑axis of a chart to use a fixed major unit of 10, overriding the automatic interval.
// Use Cases: Produce financial column charts where every Y‑axis tick marks increase by exactly 10, ensuring visual consistency across multiple reports. | Generate a series of worksheets with charts that share the same major unit, making it easy to compare data scales side‑by‑side. | Export Excel charts that comply with publishing standards requiring uniform axis spacing and fixed grid intervals.
// AI Prompts: Provide C# code to set the major unit of a line chart's value axis to 5 using Aspose.Cells. | Show how to re‑enable automatic major unit on a chart axis after it has been disabled in Aspose.Cells for .NET. | Explain how to read the current MajorUnit value from a chart's ValueAxis with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, populate sample data, add a column chart, turn off automatic major unit, set the primary Y‑axis (value axis) major unit to 10, and save the result as ChartWithMajorUnit.xlsx using Aspose.Cells for .NET (C#).
class SetMajorUnitDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 1; i <= 5; i++)
        {
            sheet.Cells[$"A{i + 1}"].PutValue("Item " + i);
            sheet.Cells[$"B{i + 1}"].PutValue(i * 12); // example values
        }

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B6", true);
        chart.NSeries.CategoryData = "A2:A6";

        // Disable automatic major unit and set it to 10 on the primary Y axis
        chart.ValueAxis.IsAutomaticMajorUnit = false;
        chart.ValueAxis.MajorUnit = 10;

        // Save the workbook with the configured chart
        workbook.Save("ChartWithMajorUnit.xlsx");
    }
}
