// Title: Create a column chart with a secondary Y‑axis and plot a specific series on it using Aspose.Cells for .NET
// AI Prompts: Generate a workbook, add sample data, create a column chart, and assign the second data series to render on a secondary Y‑axis with Aspose.Cells for .NET. | Configure the secondary axis title, minimum, maximum, and major unit for a chart built with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# add secondary Y axis to column chart | plot series on secondary axis using Aspose.Cells .NET example | set secondary value axis range and title in Aspose.Cells chart | how to display two data series with different scales in Aspose.Cells chart | Aspose.Cells chart dual Y axis configuration C#
// Tags: Aspose.Cells column chart dual axis | Aspose.Cells assign series to second axis | Aspose.Cells customize secondary axis title | Aspose.Cells set secondary axis range C# | Aspose.Cells plot series on secondary axis

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates creating a workbook, populating data, adding a column chart, rendering the second series on a secondary Y‑axis, customizing the secondary axis (title, min/max, major unit), and saving the file as ChartWithSecondaryYAxis.xlsx.
class AddSecondaryYAxisDemo
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

        sheet.Cells["B1"].PutValue("Series 1");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["B4"].PutValue(300);

        sheet.Cells["C1"].PutValue("Series 2");
        sheet.Cells["C2"].PutValue(5000);
        sheet.Cells["C3"].PutValue(3000);
        sheet.Cells["C4"].PutValue(1000);

        // Add a column chart
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIdx];

        // Add two series to the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Render the second series on the secondary Y‑axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Optional: customize the secondary Y‑axis
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.Title.Text = "Secondary Axis";
        secondaryAxis.MinValue = 0;
        secondaryAxis.MaxValue = 6000;
        secondaryAxis.MajorUnit = 1000;
        secondaryAxis.IsVisible = true;

        // Save the workbook
        workbook.Save("ChartWithSecondaryYAxis.xlsx");
    }
}
