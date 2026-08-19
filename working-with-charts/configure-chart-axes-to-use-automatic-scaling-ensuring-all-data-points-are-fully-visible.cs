// Title: Automatic Axis Scaling for Charts in Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart, and enables automatic minimum, maximum, major, minor units and tick‑label spacing for both value and category axes. The chart is recalculated and saved, ensuring every data point is fully visible without manual range settings.
// Keywords: Aspose.Cells chart auto scaling | C# automatic axis limits | Excel chart automatic min max Aspose | value axis auto scaling .NET | category axis auto scaling Aspose.Cells | chart.Calculate() Aspose | auto major unit Aspose.Cells | auto minor unit chart
// Common Searches: Aspose.Cells set chart axis to automatic | auto scaling for chart axes .NET | how to enable auto min max on Aspose chart | automatic tick label spacing Aspose.Cells | recalculate chart after axis changes Aspose
// Developer Intent: Apply Aspose.Cells properties that let Excel determine axis ranges and tick intervals automatically, so the chart adapts to any data set.
// Use Cases: Generate a column chart where axis limits adjust to dynamic data without hard‑coding values. | Create reports that automatically resize axes when new rows are added to the source range. | Ensure visual consistency across multiple workbooks by relying on Excel's built‑in scaling logic.
// AI Prompts: Show C# code to enable automatic scaling for both axes of a line chart using Aspose.Cells. | How do I toggle automatic major and minor units on a chart's value axis in Aspose.Cells for .NET? | Explain the effect of chart.Calculate() after setting IsAutomatic* properties on chart axes.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAutoScaling
{
    // Creates a workbook, adds a column chart, and enables automatic minimum, maximum, major, minor units and tick‑label spacing for both value and category axes. The chart is recalculated and saved, ensuring every data point is fully visible without manual range settings.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(25);
            worksheet.Cells["B4"].PutValue(40);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the value axis to use automatic scaling
            Axis valueAxis = chart.ValueAxis;
            valueAxis.IsAutomaticMinValue = true;   // Let Excel determine the minimum value
            valueAxis.IsAutomaticMaxValue = true;   // Let Excel determine the maximum value
            valueAxis.IsAutomaticMajorUnit = true;  // Let Excel determine the major unit
            valueAxis.IsAutomaticMinorUnit = true;  // Let Excel determine the minor unit
            valueAxis.IsAutoTickLabelSpacing = true; // Automatic tick label spacing

            // Configure the category axis to use automatic scaling as well
            Axis categoryAxis = chart.CategoryAxis;
            categoryAxis.IsAutomaticMinValue = true;
            categoryAxis.IsAutomaticMaxValue = true;
            categoryAxis.IsAutomaticMajorUnit = true;
            categoryAxis.IsAutomaticMinorUnit = true;
            categoryAxis.IsAutoTickLabelSpacing = true;

            // Recalculate the chart to apply automatic positioning and scaling
            chart.Calculate();

            // Save the workbook with the configured chart
            workbook.Save("ChartWithAutomaticScaling.xlsx", SaveFormat.Xlsx);
        }
    }
}
