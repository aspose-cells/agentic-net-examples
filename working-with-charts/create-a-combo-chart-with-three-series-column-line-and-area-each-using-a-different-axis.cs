// Title: Aspose.Cells for .NET C# – Create a Combo Chart (Column, Line, Area) with Primary and Secondary Axes
// Description: This C# example shows how to build an Excel workbook with Aspose.Cells, populate quarterly data, add a combo chart, set three series (column, line, area) each with its own chart type, plot the line series on a secondary value axis, give the chart a title, and save the file as ComboChart.xlsx.
// Keywords: Aspose.Cells combo chart C# | column line area chart .NET | secondary axis Aspose.Cells | combo chart example Aspose.Cells | multiple series chart Aspose.Cells | Excel chart with different series types | C# Aspose.Cells sample code | chart axis customization Aspose
// Common Searches: Aspose.Cells create combo chart C# | how to add line series on secondary axis Aspose.Cells | column line area chart example .NET | set individual series type Aspose.Cells chart | combo chart with multiple axes Aspose.Cells
// Developer Intent: Generate an Excel file that contains a combo chart combining column, line, and area series, with the line series displayed on a secondary axis, using Aspose.Cells for .NET in C#.
// Use Cases: Quarterly sales report: columns for sales, line for trend, area for cumulative totals. | Financial dashboard: revenue columns, profit‑margin line on secondary axis, expense distribution as area. | Project status overview: completed tasks as columns, progress percentage as line on secondary axis, resource allocation as area.
// AI Prompts: Write C# code with Aspose.Cells to add a combo chart that includes column, line, and area series, and place the line series on a secondary axis. | Explain how to assign different chart types to individual series in an Aspose.Cells combo chart and control which axis each series uses. | Provide steps to customize titles, axis labels, colors, and markers for a column‑line‑area combo chart created with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example shows how to build an Excel workbook with Aspose.Cells, populate quarterly data, add a combo chart, set three series (column, line, area) each with its own chart type, plot the line series on a secondary value axis, give the chart a title, and save the file as ComboChart.xlsx.
class ComboChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        // Column A: Categories
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["A5"].PutValue("Q4");

        // Column B: Column series values
        sheet.Cells["B1"].PutValue("ColumnSeries");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(40);

        // Column C: Line series values
        sheet.Cells["C1"].PutValue("LineSeries");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);
        sheet.Cells["C5"].PutValue(45);

        // Column D: Area series values
        sheet.Cells["D1"].PutValue("AreaSeries");
        sheet.Cells["D2"].PutValue(5);
        sheet.Cells["D3"].PutValue(15);
        sheet.Cells["D4"].PutValue(25);
        sheet.Cells["D5"].PutValue(35);

        // Add a combo chart (base type Column) to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the category (X) axis data
        chart.NSeries.CategoryData = "A2:A5";

        // Add three series to the chart
        // Series 0 – Column
        chart.NSeries.Add("B2:B5", true);
        // Series 1 – Line
        chart.NSeries.Add("C2:C5", true);
        // Series 2 – Area
        chart.NSeries.Add("D2:D5", true);

        // Configure each series type
        chart.NSeries[0].Type = ChartType.Column; // Column series (primary axis)
        chart.NSeries[1].Type = ChartType.Line;   // Line series
        chart.NSeries[2].Type = ChartType.Area;   // Area series

        // Assign axes:
        // Column series uses the primary value axis by default.
        // Line series plotted on the secondary value axis.
        chart.NSeries[1].PlotOnSecondAxis = true;
        // Area series remains on the primary axis (different from line series).

        // Optional: give the chart a title
        chart.Title.Text = "Combo Chart: Column, Line, Area";

        // Save the workbook
        workbook.Save("ComboChart.xlsx");
    }
}
