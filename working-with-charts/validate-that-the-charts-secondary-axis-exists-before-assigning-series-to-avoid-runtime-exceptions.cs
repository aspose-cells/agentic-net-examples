// Title: Check for a secondary value axis before setting PlotOnSecondAxis on a series in an Aspose.Cells column chart (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to determine if a chart has a secondary value axis and only then sets the PlotOnSecondAxis property for a specific series. | Create an example that builds a column chart with two data series, adds a secondary axis if needed, and safely assigns the second series to that axis.
// Common Searches: Aspose.Cells C# how to conditionally plot a series on secondary axis | prevent runtime error when using PlotOnSecondAxis in Aspose.Cells chart | check chart.HasAxis for secondary value axis before assigning series Aspose.Cells | C# example of adding secondary axis to column chart with Aspose.Cells | validate secondary axis existence Aspose.Cells before PlotOnSecondAxis
// Tags: chart.HasAxis secondary value axis check | PlotOnSecondAxis conditional assignment Aspose.Cells | column chart secondary axis handling C# | Aspose.Cells chart series secondary axis validation | prevent PlotOnSecondAxis exception Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills it with category and two data series, adds a column chart, checks whether a secondary value axis exists using chart.HasAxis, and only sets PlotOnSecondAxis for the second series when the axis is present, then saves the file as ValidateSecondaryAxis.xlsx.
class ValidateSecondaryAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for categories and two series
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(100);
        sheet.Cells["C3"].PutValue(200);
        sheet.Cells["C4"].PutValue(300);

        // Add a column chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];

        // Add two series to the chart
        chart.NSeries.Add("B2:B4", true); // first series
        chart.NSeries.Add("C2:C4", true); // second series
        chart.NSeries.CategoryData = "A2:A4";

        // Validate that the secondary value axis exists before assigning a series to it
        bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);
        if (hasSecondaryValueAxis)
        {
            // Safe to plot the second series on the secondary axis
            chart.NSeries[1].PlotOnSecondAxis = true;
        }
        else
        {
            // Secondary axis does not exist; handle the situation as needed
            Console.WriteLine("Secondary value axis not present. Skipping PlotOnSecondAxis assignment.");
        }

        // Save the workbook to a file
        workbook.Save("ValidateSecondaryAxis.xlsx");
    }
}
