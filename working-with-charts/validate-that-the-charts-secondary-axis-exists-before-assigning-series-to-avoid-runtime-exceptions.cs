// Title: Validate Secondary Axis Before Plotting Series in an Aspose.Cells Column Chart (C#)
// Description: Creates a workbook, adds primary and secondary data, inserts a column chart, checks if the secondary value axis exists with chart.HasAxis, makes the axis visible when missing, safely assigns the second series to the secondary axis, customizes axis titles and ranges, and saves the file—preventing runtime exceptions.
// Keywords: Aspose.Cells | C# chart secondary axis | validate secondary axis | chart.HasAxis | plot series on second axis | runtime exception prevention | column chart Aspose.Cells | global developers | US C# developers | European .NET charting
// Common Searches: how to check for secondary axis in Aspose.Cells chart | Aspose.Cells C# create secondary value axis if not present | prevent error when assigning series to second axis Aspose.Cells | chart.HasAxis usage example C# | add secondary axis to column chart Aspose.Cells
// Developer Intent: Ensure a secondary value axis is present before assigning a series to avoid runtime errors.
// Use Cases: Display data sets with different scales on the same column chart using a secondary axis. | Programmatically enable a secondary axis only when required, keeping the workbook lightweight. | Customize secondary axis properties (title, min/max, major unit) after confirming its existence.
// AI Prompts: Write C# code with Aspose.Cells that verifies a secondary value axis and creates it if absent before plotting a series. | Explain the interaction between chart.HasAxis and chart.SecondValueAxis.IsVisible for safe secondary axis creation. | Provide a step‑by‑step tutorial to assign a series to the secondary axis without triggering a runtime exception.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds primary and secondary data, inserts a column chart, checks if the secondary value axis exists with chart.HasAxis, makes the axis visible when missing, safely assigns the second series to the secondary axis, customizes axis titles and ranges, and saves the file—preventing runtime exceptions.
class ValidateSecondaryAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        worksheet.Cells["B1"].PutValue("Primary");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        worksheet.Cells["C1"].PutValue("Secondary");
        worksheet.Cells["C2"].PutValue(500);
        worksheet.Cells["C3"].PutValue(300);
        worksheet.Cells["C4"].PutValue(100);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Add two series: primary and secondary
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Verify that the secondary value axis exists
        bool hasSecondaryAxis = chart.HasAxis(AxisType.Value, false);
        if (!hasSecondaryAxis)
        {
            // Making the secondary axis visible will create it if it does not exist
            chart.SecondValueAxis.IsVisible = true;
        }

        // Safely assign the second series to the secondary axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Optional: customize the secondary axis appearance
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.Title.Text = "Secondary Axis";
        secondaryAxis.MinValue = 0;
        secondaryAxis.MaxValue = 600;
        secondaryAxis.MajorUnit = 100;

        // Save the workbook
        workbook.Save("ChartWithValidatedSecondaryAxis.xlsx");
    }
}
