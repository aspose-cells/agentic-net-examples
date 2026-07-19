// Title: Create a Stacked Column Chart with Three Series and Individual Colors using Aspose.Cells for .NET (C#)
// Description: This example shows how to build a new Workbook, fill cells A1:D5 with category labels and three data series, add a ColumnStacked chart, bind the series to range B2:D5, set the category axis to A2:A5, assign a unique foreground color to each series via the Area.ForegroundColor property, and save the result as StackedColumnChart.xlsx.
// Keywords: Aspose.Cells | C# | .NET | stacked column chart | custom series color | ChartType.ColumnStacked | Area.ForegroundColor | Excel chart automation | multiple series chart | programmatic chart styling
// Common Searches: Aspose.Cells set individual series colors stacked column chart C# | how to change series color in Aspose.Cells chart | create stacked column chart with three series Aspose.Cells .NET | customize chart series colors programmatically Aspose.Cells | add multiple series to Excel chart using Aspose.Cells
// Developer Intent: Generate a stacked column chart with three data series and assign a distinct color to each series via code.
// Use Cases: Quarterly sales dashboard where each product line appears as a uniquely colored segment in a stacked column chart. | Financial report that visualizes expense categories with brand‑specific colors for clear comparison. | Presentation‑ready chart highlighting three performance metrics, each distinguished by a custom color.
// AI Prompts: Provide C# code to add data labels to each series of the stacked column chart created with Aspose.Cells. | Show how to load series colors from a JSON configuration file and apply them to the chart after it is generated. | Explain how to export the stacked column chart as a high‑resolution PNG while preserving the custom series colors.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to build a new Workbook, fill cells A1:D5 with category labels and three data series, add a ColumnStacked chart, bind the series to range B2:D5, set the category axis to A2:A5, assign a unique foreground color to each series via the Area.ForegroundColor property, and save the result as StackedColumnChart.xlsx.
class StackedColumnChartExample
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

        // Column B: Series 1 values
        sheet.Cells["B1"].PutValue("Series 1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(40);

        // Column C: Series 2 values
        sheet.Cells["C1"].PutValue("Series 2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);
        sheet.Cells["C5"].PutValue(45);

        // Column D: Series 3 values
        sheet.Cells["D1"].PutValue("Series 3");
        sheet.Cells["D2"].PutValue(20);
        sheet.Cells["D3"].PutValue(30);
        sheet.Cells["D4"].PutValue(40);
        sheet.Cells["D5"].PutValue(50);

        // Add a stacked column chart
        int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Add three series (B2:D5) vertically; each column becomes a series
        chart.NSeries.Add("B2:D5", true);
        // Set category (X‑axis) data
        chart.NSeries.CategoryData = "A2:A5";

        // Customize each series color individually
        chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(79, 129, 189);   // Series 1 color
        chart.NSeries[1].Area.ForegroundColor = Color.FromArgb(192, 80, 77);   // Series 2 color
        chart.NSeries[2].Area.ForegroundColor = Color.FromArgb(155, 187, 89);  // Series 3 color

        // Save the workbook
        workbook.Save("StackedColumnChart.xlsx");
    }
}
