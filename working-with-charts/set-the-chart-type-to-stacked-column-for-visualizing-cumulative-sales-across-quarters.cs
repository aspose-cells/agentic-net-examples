// Title: Generate a stacked column chart for cumulative quarterly sales using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that adds a ColumnStacked chart to display quarterly sales totals. | Show how to assign the data range and category axis for a stacked column chart and set a custom chart title. | Provide the steps to save the workbook containing the stacked column chart as an XLSX file.
// Common Searches: asp.net how to add a stacked column chart with Aspose.Cells C# | c# Aspose.Cells example for cumulative sales per quarter chart | setting data series and categories for ColumnStacked chart in Aspose.Cells | saving workbook with chart as xlsx using Aspose.Cells C#
// Tags: Aspose.Cells stacked column visualization | Aspose.Cells chart data series setup | Aspose.Cells XLSX chart export | Aspose.Cells chart title setting | Aspose.Cells cumulative sales graph

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates creating a workbook, populating quarterly sales data, adding a ColumnStacked chart, configuring its data series and categories, applying a custom title, and saving the result as StackedColumnChart.xlsx.
class StackedColumnChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample cumulative sales data across quarters
        sheet.Cells["A1"].PutValue("Quarter");
        sheet.Cells["B1"].PutValue("Product A");
        sheet.Cells["C1"].PutValue("Product B");
        sheet.Cells["D1"].PutValue("Product C");

        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["C2"].PutValue(150);
        sheet.Cells["D2"].PutValue(100);

        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B3"].PutValue(130);
        sheet.Cells["C3"].PutValue(160);
        sheet.Cells["D3"].PutValue(110);

        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["B4"].PutValue(140);
        sheet.Cells["C4"].PutValue(170);
        sheet.Cells["D4"].PutValue(120);

        sheet.Cells["A5"].PutValue("Q4");
        sheet.Cells["B5"].PutValue(150);
        sheet.Cells["C5"].PutValue(180);
        sheet.Cells["D5"].PutValue(130);

        // Add a stacked column chart (ChartType.ColumnStacked)
        int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 7, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series (columns B to D) and categories (column A)
        chart.NSeries.Add("B2:D5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Optional: set a chart title
        chart.Title.Text = "Cumulative Sales by Quarter";

        // Save the workbook with the chart
        workbook.Save("StackedColumnChart.xlsx", SaveFormat.Xlsx);
    }
}
