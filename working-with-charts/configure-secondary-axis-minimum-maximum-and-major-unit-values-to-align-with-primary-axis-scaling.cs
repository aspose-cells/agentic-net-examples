// Title: Aspose.Cells C# – Align Secondary Axis Min, Max, and Major Unit with Primary Axis in a Column Chart
// Description: Learn how to disable automatic scaling and set identical MinValue, MaxValue, and MajorUnit for the secondary value axis as the primary axis in an Aspose.Cells column chart. The example creates a workbook, adds a chart with two series of different ranges, configures both axes manually, and saves the file.
// Keywords: Aspose.Cells secondary axis | C# chart axis scaling | set secondary axis min value Aspose | align secondary and primary axis Aspose.Cells | Excel column chart secondary value axis | Aspose.Cells chart axis properties | .NET chart axis configuration
// Common Searches: Aspose.Cells set secondary axis minimum C# | how to match secondary axis max value with primary axis Aspose | C# Aspose.Cells chart major unit secondary axis | disable automatic axis scaling Aspose.Cells | example aligning secondary axis scaling Aspose.Cells
// Developer Intent: Manually set the secondary value axis limits and major unit to be identical to the primary axis in an Aspose.Cells chart.
// Use Cases: Display two data series with vastly different magnitudes while keeping a unified visual scale. | Create Excel reports where both primary and secondary axes must share the same range for consistency. | Generate charts that render correctly across Excel, Google Sheets, and other viewers by fixing axis values.
// AI Prompts: Show C# code using Aspose.Cells to set secondary axis MinValue, MaxValue, and MajorUnit equal to the primary axis for any chart type. | Explain which Aspose.Cells properties control automatic axis scaling and how to turn them off for both axes. | Provide a step‑by‑step tutorial to align secondary axis scaling with the primary axis, including sample data, chart creation, and workbook saving.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryAxisDemo
{
    // Learn how to disable automatic scaling and set identical MinValue, MaxValue, and MajorUnit for the secondary value axis as the primary axis in an Aspose.Cells column chart. The example creates a workbook, adds a chart with two series of different ranges, configures both axes manually, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Category");
            cells["A2"].PutValue("A");
            cells["A3"].PutValue("B");
            cells["A4"].PutValue("C");

            // Primary series values (smaller range)
            cells["B1"].PutValue("Primary Series");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);

            // Secondary series values (larger range)
            cells["C1"].PutValue("Secondary Series");
            cells["C2"].PutValue(1000);
            cells["C3"].PutValue(2000);
            cells["C4"].PutValue(3000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add primary series and bind category data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add secondary series
            chart.NSeries.Add("C2:C4", true);
            // Plot the second series on the secondary value axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // ----- Primary (value) axis configuration -----
            Axis primaryAxis = chart.ValueAxis;
            primaryAxis.IsAutomaticMinValue = false;
            primaryAxis.IsAutomaticMaxValue = false;
            primaryAxis.IsAutomaticMajorUnit = false;

            primaryAxis.MinValue = 0;      // Minimum value
            primaryAxis.MaxValue = 40;     // Maximum value
            primaryAxis.MajorUnit = 10;    // Major unit interval

            // ----- Secondary (value) axis configuration -----
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.IsAutomaticMinValue = false;
            secondaryAxis.IsAutomaticMaxValue = false;
            secondaryAxis.IsAutomaticMajorUnit = false;

            // Align secondary axis scaling with primary axis
            secondaryAxis.MinValue = 0;
            secondaryAxis.MaxValue = 40;
            secondaryAxis.MajorUnit = 10;

            // Save the workbook
            workbook.Save("SecondaryAxisAligned.xlsx");
        }
    }
}
