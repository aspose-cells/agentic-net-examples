// Title: C# – Assign a Line Series to the Secondary Vertical Axis with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a line chart with two data series, and move the second series to the secondary vertical axis by setting its AxisGroup (or PlotOnSecondAxis) property, then save the file as an Excel workbook.
// Keywords: Aspose.Cells C# line chart secondary axis | PlotOnSecondAxis example | AxisGroup property Aspose.Cells | assign series to secondary vertical axis | C# Excel chart programming | Aspose.Cells chart series axis
// Common Searches: Aspose.Cells plot series on secondary axis C# | how to use AxisGroup=2 in Aspose.Cells chart | C# line chart secondary vertical axis Aspose | move chart series to secondary axis programmatically | Aspose.Cells secondary axis tutorial
// Developer Intent: Place the second line series on the chart’s secondary vertical axis.
// Use Cases: Display metrics with different scales side‑by‑side in a single line chart. | Generate Excel reports where one data series requires its own axis range. | Update existing workbooks to shift a specific series to a secondary axis without recreating the chart.
// AI Prompts: Generate C# code using Aspose.Cells that adds a line chart with two series and assigns the second series to the secondary vertical axis. | Explain how the AxisGroup property (or PlotOnSecondAxis) controls primary vs. secondary axis placement in Aspose.Cells charts. | Show how to programmatically switch a chart series from the primary to the secondary axis in an existing Excel file with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a line chart with two data series, and move the second series to the secondary vertical axis by setting its AxisGroup (or PlotOnSecondAxis) property, then save the file as an Excel workbook.
class AssignSeriesToSecondaryAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Jan");
        worksheet.Cells["A3"].PutValue("Feb");
        worksheet.Cells["A4"].PutValue("Mar");

        worksheet.Cells["B1"].PutValue("Primary");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        worksheet.Cells["C1"].PutValue("Secondary");
        worksheet.Cells["C2"].PutValue(100);
        worksheet.Cells["C3"].PutValue(200);
        worksheet.Cells["C4"].PutValue(300);

        // Add a line chart
        int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Add two series: primary and secondary
        chart.NSeries.Add("B2:B4", true);   // Primary series
        chart.NSeries.Add("C2:C4", true);   // Secondary series
        chart.NSeries.CategoryData = "A2:A4";

        // Assign the second series to the secondary vertical axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Save the workbook
        workbook.Save("LineSeriesSecondaryAxis.xlsx");
    }
}
