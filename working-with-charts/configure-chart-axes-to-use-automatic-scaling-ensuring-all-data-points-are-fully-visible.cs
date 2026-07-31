// Title: Auto‑scale chart axes in Aspose.Cells for .NET (C#) – show all data points
// Description: C# example that creates a workbook, adds a column chart, and configures both value and category axes to use automatic minimum, maximum, major and minor units. The plot area is set to auto‑position, the chart is recalculated, and the file is saved as an XLSX.
// Keywords: Aspose.Cells chart auto scaling | C# automatic axis min max | Aspose.Cells value axis auto | category axis automatic scaling | chart plot area auto position | Aspose.Cells chart.Calculate | auto adjust chart axes .NET
// Common Searches: Aspose.Cells set chart axis to automatic scaling C# | how to enable auto min max for chart axes in Aspose.Cells | auto major unit and minor unit Aspose.Cells chart | C# Aspose.Cells auto‑position plot area | recalculate chart after axis changes Aspose.Cells
// Developer Intent: Configure a chart so Excel automatically determines axis ranges and units, guaranteeing every data point is displayed without manual settings.
// Use Cases: Generate sales or KPI charts where axis limits adapt to changing data. | Build dynamic reports with multiple charts that self‑adjust when data updates. | Create templates that automatically position the plot area for optimal layout before saving.
// AI Prompts: Show C# code to enable automatic scaling for both value and category axes in an Aspose.Cells chart. | Explain why calling chart.Calculate() is required after setting axis auto properties in Aspose.Cells. | Provide an Aspose.Cells example that auto‑positions the plot area and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartAxisAutoScalingDemo
{
    // C# example that creates a workbook, adds a column chart, and configures both value and category axes to use automatic minimum, maximum, major and minor units. The plot area is set to auto‑position, the chart is recalculated, and the file is saved as an XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
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

            // Ensure the plot area positioning is automatic
            chart.PlotArea.SetPositionAuto();

            // Recalculate the chart so that automatic settings take effect
            chart.Calculate();

            // Save the workbook with the configured chart
            workbook.Save("ChartAxisAutoScalingDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
