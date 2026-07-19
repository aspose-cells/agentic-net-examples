// Title: Automatic Axis Scaling for Column Charts with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a column chart from worksheet data, enable automatic minimum, maximum, major and minor units for both value and category axes, recalculate the chart, and save the workbook. The axes adjust dynamically so every data point stays visible, making the solution suitable for reports that change over time.
// Keywords: Aspose.Cells chart auto scaling | C# automatic axis limits | value axis IsAutomaticMinValue | category axis IsAutomaticMaxValue | chart.Calculate Aspose.Cells | dynamic Excel chart scaling | Aspose.Cells .NET chart example
// Common Searches: Aspose.Cells set chart axes to auto scale C# | How to enable automatic min/max for chart axes in .NET | Aspose.Cells automatic major unit for column chart | C# chart axis auto scaling Aspose.Cells tutorial | Make Excel chart axes adjust automatically with Aspose
// Developer Intent: Configure both value and category axes to scale automatically so the chart always displays the full data range.
// Use Cases: Generating dashboards where data ranges vary and manual axis tweaks are impractical. | Automating financial or sales reports that require charts to adapt to new values each run. | Building reusable Excel templates that preserve proper axis scaling after data updates.
// AI Prompts: Show C# code that sets IsAutomaticMinValue, IsAutomaticMaxValue, IsAutomaticMajorUnit, and IsAutomaticMinorUnit for a chart's value and category axes using Aspose.Cells. | Provide an example that creates a line chart, applies automatic scaling to both axes, calls chart.Calculate, and saves the workbook. | Explain why calling chart.Calculate is necessary after configuring automatic axis properties in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAutoScalingDemo
{
    // Demonstrates how to create a column chart from worksheet data, enable automatic minimum, maximum, major and minor units for both value and category axes, recalculate the chart, and save the workbook. The axes adjust dynamically so every data point stays visible, making the solution suitable for reports that change over time.
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
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["B4"].PutValue(40);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the value axis to use automatic scaling
            Axis valueAxis = chart.ValueAxis;
            valueAxis.IsAutomaticMinValue = true;   // Let Excel determine the minimum value
            valueAxis.IsAutomaticMaxValue = true;   // Let Excel determine the maximum value
            valueAxis.IsAutomaticMajorUnit = true;  // Let Excel determine the major unit
            valueAxis.IsAutomaticMinorUnit = true;  // Let Excel determine the minor unit

            // Configure the category axis to use automatic scaling as well
            Axis categoryAxis = chart.CategoryAxis;
            categoryAxis.IsAutomaticMinValue = true;
            categoryAxis.IsAutomaticMaxValue = true;
            categoryAxis.IsAutomaticMajorUnit = true;
            categoryAxis.IsAutomaticMinorUnit = true;

            // Recalculate the chart to apply automatic positioning and scaling
            chart.Calculate();

            // Save the workbook with the configured chart
            workbook.Save("ChartAutoScalingDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
