// Title: Create a stacked column chart with three series and assign red, green, and blue colors to each series using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that builds a ColumnStacked chart with three data series, binds category labels from column A, and sets the first series fill to red, the second to green, and the third to blue using Aspose.Cells. | Show how to add multiple series to a stacked column chart from worksheet ranges and customize each series' foreground color programmatically in Aspose.Cells. | Provide a complete example that saves the workbook containing the customized stacked column chart to an .xlsx file with Aspose.Cells.
// Common Searches: asp.net aspose.cells create stacked column chart with three series custom colors | c# set individual series colors in a ColumnStacked chart using Aspose.Cells | how to bind category data and multiple series to a stacked column chart in Aspose.Cells | asp.net example for customizing series fill color in Aspose.Cells chart | save workbook with stacked column chart as .xlsx using Aspose.Cells C#
// Tags: stacked column chart creation Aspose.Cells C# | custom series fill color Aspose.Cells | bind worksheet range to chart series Aspose.Cells | add multiple series to ColumnStacked chart Aspose.Cells | save workbook as .xlsx Aspose.Cells | set series foreground color Aspose.Cells chart

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, fills columns A‑D with category labels and three data series, adds a ColumnStacked chart, binds the series and category ranges, sets the first series fill to red, second to green, third to blue, and saves the file as StackedColumnChartWithCustomColors.xlsx.
class StackedColumnChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        // Column A – Category labels
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");

        // Column B – Series 1 values
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Column C – Series 2 values
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Column D – Series 3 values
        sheet.Cells["D1"].PutValue("Series3");
        sheet.Cells["D2"].PutValue(12);
        sheet.Cells["D3"].PutValue(22);
        sheet.Cells["D4"].PutValue(32);

        // Add a stacked column chart
        int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 6, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add the three series (by column) and set category data
        // The first Add call adds all three series because the range includes B:D rows 2-4
        chart.NSeries.Add("B2:D4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Customize each series color individually
        // Series 0 – red
        chart.NSeries[0].Area.ForegroundColor = Color.Red;
        // Series 1 – green
        chart.NSeries[1].Area.ForegroundColor = Color.Green;
        // Series 2 – blue
        chart.NSeries[2].Area.ForegroundColor = Color.Blue;

        // Save the workbook
        workbook.Save("StackedColumnChartWithCustomColors.xlsx");
    }
}
