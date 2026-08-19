// Title: C# Aspose.Cells: Set Y‑Axis Scale 0‑1000 on a Bar (Column) Chart
// Description: This example creates a workbook, adds a column chart with sample data, disables automatic axis scaling, and explicitly sets the Y‑axis (value axis) minimum to 0, maximum to 1000, and major unit to 200 before saving the file as BarChart_YAxis_0_to_1000.xlsx.
// Keywords: Aspose.Cells Y axis range C# | set chart axis minimum maximum Aspose.Cells | C# bar chart axis scaling | Aspose.Cells column chart Y axis | custom Y axis limits .NET | Excel chart value axis Aspose | disable automatic axis scaling Aspose.Cells
// Common Searches: Aspose.Cells set Y axis minimum and maximum C# | how to fix Y axis range 0 to 1000 in Aspose.Cells chart | C# Aspose.Cells change chart major unit | disable automatic Y axis scaling Aspose.Cells .NET | set value axis limits for column chart Aspose
// Developer Intent: Programmatically fix the Y‑axis scale of a bar/column chart to a defined range (0‑1000) using Aspose.Cells for .NET.
// Use Cases: Standardize sales dashboards so every bar chart shares a 0‑1000 Y‑axis for easy comparison. | Create performance reports that require a fixed axis to meet corporate presentation guidelines. | Export analytical Excel files where automatic rescaling could mislead viewers.
// AI Prompts: Show C# code to set Y‑axis minimum 0, maximum 1000, and major unit 200 for a column chart with Aspose.Cells. | How do I disable automatic Y‑axis scaling and apply a custom range to all charts in a workbook using Aspose.Cells .NET? | Explain step‑by‑step how to configure the value axis of a bar chart in Aspose.Cells for C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsYAxisScaleExample
{
    // This example creates a workbook, adds a column chart with sample data, disables automatic axis scaling, and explicitly sets the Y‑axis (value axis) minimum to 0, maximum to 1000, and major unit to 200 before saving the file as BarChart_YAxis_0_to_1000.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the bar chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(200);
            worksheet.Cells["B3"].PutValue(600);
            worksheet.Cells["B4"].PutValue(950);

            // Add a bar chart (Column chart works as a vertical bar chart)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the Y‑axis (value axis) to display from 0 to 1000
            Axis valueAxis = chart.ValueAxis;

            // Turn off automatic min/max calculation
            valueAxis.IsAutomaticMinValue = false;
            valueAxis.IsAutomaticMaxValue = false;

            // Set explicit minimum and maximum values
            valueAxis.MinValue = 0;      // Minimum value
            valueAxis.MaxValue = 1000;   // Maximum value

            // Optional: set major unit for clearer grid lines
            valueAxis.MajorUnit = 200;

            // Save the workbook to a file
            workbook.Save("BarChart_YAxis_0_to_1000.xlsx");
        }
    }
}
