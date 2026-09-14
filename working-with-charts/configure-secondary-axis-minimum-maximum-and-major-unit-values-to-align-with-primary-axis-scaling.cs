// Title: How to set matching minimum, maximum, and major unit values for primary and secondary Y‑axes in an Aspose.Cells column chart (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a column chart with dual Y‑axes and manually set identical MinValue, MaxValue, and MajorUnit for both the primary and secondary value axes. | Show how to disable automatic scaling and copy axis scaling properties from the primary ValueAxis to the SecondValueAxis in an Aspose.Cells workbook. | Provide a complete example that adds sales and profit series, plots profit on the secondary axis, and aligns the secondary axis bounds with the primary axis using Aspose.Cells.
// Common Searches: Aspose.Cells C# set secondary axis min and max same as primary axis | how to copy axis scaling from primary to secondary value axis in Aspose.Cells chart | C# Aspose.Cells dual axis column chart manual scaling | disable automatic axis scaling Aspose.Cells chart C# | set major unit for secondary Y axis Aspose.Cells
// Tags: Aspose.Cells set axis scaling programmatically | C# Aspose.Cells secondary value axis configuration | dual-axis column chart axis bounds Aspose.Cells | manual axis min max major unit Aspose.Cells | copy primary axis properties to secondary axis Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryAxisScaling
{
    // The example creates a workbook, adds sales and profit data, builds a column chart with two series, plots the profit series on the secondary Y‑axis, disables automatic scaling, sets the primary axis MinValue, MaxValue, and MajorUnit to 0, 100, and 20, copies those scaling values to the secondary axis, assigns titles to both axes, and saves the file as ChartWithAlignedSecondaryAxis.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for two series
            // Primary series (plotted on primary Y axis)
            cells["A1"].PutValue("Category");
            cells["A2"].PutValue("Jan");
            cells["A3"].PutValue("Feb");
            cells["A4"].PutValue("Mar");
            cells["B1"].PutValue("Sales");
            cells["B2"].PutValue(40);
            cells["B3"].PutValue(60);
            cells["B4"].PutValue(80);

            // Secondary series (plotted on secondary Y axis)
            cells["C1"].PutValue("Profit");
            cells["C2"].PutValue(2000);
            cells["C3"].PutValue(3000);
            cells["C4"].PutValue(4000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the two series to the chart
            chart.NSeries.Add("B2:B4", true);          // Primary series
            chart.NSeries.Add("C2:C4", true);          // Secondary series
            chart.NSeries.CategoryData = "A2:A4";

            // Plot the second series on the secondary value axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Configure primary value axis scaling
            Axis primaryAxis = chart.ValueAxis;
            primaryAxis.IsAutomaticMinValue = false;
            primaryAxis.IsAutomaticMaxValue = false;
            primaryAxis.IsAutomaticMajorUnit = false;

            primaryAxis.MinValue = 0;      // Minimum
            primaryAxis.MaxValue = 100;    // Maximum
            primaryAxis.MajorUnit = 20;    // Major unit

            // Align secondary axis scaling with primary axis
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.IsAutomaticMinValue = false;
            secondaryAxis.IsAutomaticMaxValue = false;
            secondaryAxis.IsAutomaticMajorUnit = false;

            // Copy scaling values from primary axis
            secondaryAxis.MinValue = primaryAxis.MinValue;
            secondaryAxis.MaxValue = primaryAxis.MaxValue;
            secondaryAxis.MajorUnit = primaryAxis.MajorUnit;

            // Optional: give titles to axes for clarity
            primaryAxis.Title.Text = "Sales";
            secondaryAxis.Title.Text = "Profit";

            // Save the workbook
            workbook.Save("ChartWithAlignedSecondaryAxis.xlsx");
        }
    }
}
