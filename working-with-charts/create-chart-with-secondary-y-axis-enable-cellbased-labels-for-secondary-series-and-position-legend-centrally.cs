// Title: Aspose.Cells C# – Create Column Chart with Secondary Y‑Axis, Data Labels and Centered Bottom Legend
// Description: This example shows how to build an Excel workbook with Aspose.Cells, add a column chart, define primary and secondary series, plot the second series on a secondary Y‑axis, customize its range and title, enable value data labels for that series, and place a centered legend at the bottom before saving the file.
// Keywords: Aspose.Cells secondary axis chart C# | column chart with legend center Aspose | enable data labels secondary series .NET | custom secondary value axis Aspose.Cells | chart legend position bottom Aspose
// Common Searches: Aspose.Cells plot series on secondary Y axis | how to center legend at bottom in Aspose.Cells chart | enable data labels for specific series Aspose.Cells | set secondary axis min max Aspose.Cells | C# create chart with two Y axes using Aspose
// Developer Intent: Generate an Excel column chart where one series uses a secondary Y‑axis, display its values as data labels, and show a horizontally centered legend at the bottom, all via Aspose.Cells for .NET.
// Use Cases: Compare revenue (primary axis) and profit margin (secondary axis) in a single column chart. | Show temperature on the primary axis and humidity on a secondary axis with distinct scales and labels. | Create a financial dashboard that plots sales volume and expense ratio together, using a centered legend for quick reference.
// AI Prompts: Write C# code with Aspose.Cells to add a line series to the secondary Y‑axis and format its data labels as percentages. | Demonstrate how to bind data labels to cell values instead of raw numbers for a chart series in Aspose.Cells. | Explain how to change the font style of the secondary axis title and move the legend to the top‑right corner using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to build an Excel workbook with Aspose.Cells, add a column chart, define primary and secondary series, plot the second series on a secondary Y‑axis, customize its range and title, enable value data labels for that series, and place a centered legend at the bottom before saving the file.
class CreateChartWithSecondaryAxis
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Primary");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);

            sheet.Cells["C1"].PutValue("Secondary");
            sheet.Cells["C2"].PutValue(5000);
            sheet.Cells["C3"].PutValue(3000);
            sheet.Cells["C4"].PutValue(1000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add two series: primary and secondary
            chart.NSeries.Add("B2:B4", true); // primary series
            chart.NSeries.Add("C2:C4", true); // secondary series
            chart.NSeries.CategoryData = "A2:A4";

            // Plot the second series on the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Customize the secondary Y axis
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Secondary Axis";
            secondaryAxis.MinValue = 0;
            secondaryAxis.MaxValue = 6000;
            secondaryAxis.MajorUnit = 1000;

            // Enable data labels for the secondary series
            chart.NSeries[1].DataLabels.ShowValue = true;

            // Position the legend at the bottom and center it horizontally
            chart.Legend.Position = LegendPositionType.Bottom;
            chart.Legend.XRatioToChart = 0.5; // center horizontally

            // Recalculate chart layout
            chart.Calculate();

            // Save the workbook
            workbook.Save("ChartWithSecondaryAxis.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
