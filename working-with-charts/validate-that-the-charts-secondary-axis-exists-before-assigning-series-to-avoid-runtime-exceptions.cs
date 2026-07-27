// Title: Safely Assign a Series to a Chart’s Secondary Axis with Aspose.Cells for .NET
// Description: Demonstrates how to verify the presence of a secondary value axis using chart.HasAxis before setting PlotOnSecondAxis, create the axis when missing, and customize its appearance in a C# column chart.
// Keywords: Aspose.Cells secondary axis | chart.HasAxis C# | PlotOnSecondAxis example | validate secondary value axis | avoid runtime exception Aspose.Cells | .NET chart axis check | secondary axis customization
// Common Searches: Aspose.Cells check secondary axis before PlotOnSecondAxis | C# chart secondary axis validation example | how to enable secondary value axis in Aspose.Cells | prevent error when assigning series to secondary axis | chart.HasAxis usage Aspose.Cells .NET
// Developer Intent: Confirm that a secondary value axis exists and make it visible if necessary before assigning a series to it, eliminating runtime errors.
// Use Cases: Generating column charts that combine primary and secondary data series in automated reports. | Building dynamic dashboards where chart types may change (e.g., from column to pie) and secondary axes need conditional handling. | Customizing secondary axis titles, ranges, and visibility only after confirming the axis is present.
// AI Prompts: Create a C# method that receives a Chart object and a series index, checks for a secondary axis with chart.HasAxis, adds and shows the axis if absent, then sets PlotOnSecondAxis safely. | Write Aspose.Cells code to add a line chart with two data series, validate the secondary value axis, and assign the second series to it while handling chart types without a secondary axis. | Explain the behavior of chart.HasAxis for primary and secondary axes and show how to programmatically make a secondary axis visible in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryAxisCheck
{
    // Demonstrates how to verify the presence of a secondary value axis using chart.HasAxis before setting PlotOnSecondAxis, create the axis when missing, and customize its appearance in a C# column chart.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Primary");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Secondary");
            sheet.Cells["C2"].PutValue(100);
            sheet.Cells["C3"].PutValue(200);
            sheet.Cells["C4"].PutValue(300);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Add two series: primary and secondary
            chart.NSeries.Add("B2:B4", true);          // Primary series
            chart.NSeries.Add("C2:C4", true);          // Series intended for secondary axis
            chart.NSeries.CategoryData = "A2:A4";

            // Validate that a secondary value axis exists before assigning PlotOnSecondAxis
            bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);
            if (hasSecondaryValueAxis)
            {
                // The secondary axis exists – safely assign the series to it
                chart.NSeries[1].PlotOnSecondAxis = true;
            }
            else
            {
                // If the secondary axis does not exist (e.g., pie chart), optionally handle it
                // For demonstration, we make the secondary axis visible (if applicable) and then assign
                chart.SecondValueAxis.IsVisible = true;
                chart.NSeries[1].PlotOnSecondAxis = true;
            }

            // Optional: customize the secondary axis appearance
            Axis secAxis = chart.SecondValueAxis;
            secAxis.Title.Text = "Secondary Axis";
            secAxis.MinValue = 0;
            secAxis.MaxValue = 400;
            secAxis.MajorUnit = 100;

            // Save the workbook
            workbook.Save("ChartWithValidatedSecondaryAxis.xlsx");
        }
    }
}
