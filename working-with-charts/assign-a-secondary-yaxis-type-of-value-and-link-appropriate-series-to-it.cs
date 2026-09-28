// Title: Create a column chart with a secondary value axis and map a data series to it using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that builds a workbook, adds a column chart, creates a secondary value axis, and assigns the second series to that axis with Aspose.Cells. | Show how to set the title, minimum, maximum, and major unit of the secondary Y‑axis for a column chart in Aspose.Cells. | Provide an example that saves the workbook after configuring a secondary axis and verifies the chart displays both primary and secondary series correctly.
// Common Searches: how to plot a second series on a secondary y‑axis in Aspose.Cells C# column chart | Aspose.Cells set secondary value axis range and title for column chart | C# Aspose.Cells add secondary Y axis of type Value to chart | configure secondary axis scale in Aspose.Cells chart programmatically | example of secondary axis with different units in Aspose.Cells .NET
// Tags: Aspose.Cells column chart secondary value axis | C# assign series to secondary Y axis Aspose.Cells | configure secondary axis title and scale Aspose.Cells | Aspose.Cells chart secondary axis range | add secondary Y axis to chart Aspose.Cells .NET | Aspose.Cells chart series on second axis

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates creating a workbook, inserting sample data, adding a column chart, defining a secondary value axis, assigning the second data series to that axis, customizing the axis title and scale, and saving the workbook using Aspose.Cells for .NET.
class SecondaryYAxisDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data: categories and two data series
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series to the chart
        chart.NSeries.Add("B2:B4", true); // primary series
        chart.NSeries.Add("C2:C4", true); // secondary series
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary Y‑axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Configure the secondary value axis (type is Value by default)
        Axis secondaryValueAxis = chart.SecondValueAxis;
        secondaryValueAxis.Title.Text = "Secondary Values";
        secondaryValueAxis.MinValue = 0;
        secondaryValueAxis.MaxValue = 6000;
        secondaryValueAxis.MajorUnit = 1000;

        // Save the workbook
        workbook.Save("SecondaryYAxisDemo.xlsx");
    }
}
