// Title: Aspose.Cells for .NET – C# Example: Create a Stacked Column Chart for Cumulative Quarterly Sales
// Description: This C# sample builds a new workbook, fills it with quarterly sales data for three products, adds a ColumnStacked chart (quarters on the X‑axis, sales values as series), sets a title, and saves the file as StackedColumnChart.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells C# stacked column chart | ColumnStacked chart example | cumulative quarterly sales Excel | add chart to workbook Aspose.Cells | ChartType.ColumnStacked .NET | Excel chart series and category data | Aspose.Cells chart tutorial | visualize sales with stacked columns
// Common Searches: Aspose.Cells create stacked column chart C# | how to add ColumnStacked chart in .NET | cumulative sales chart using Aspose.Cells | set chart type to ColumnStacked Aspose.Cells | example of stacked column chart in Excel with C#
// Developer Intent: Generate a stacked column chart to display cumulative sales per quarter in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Produce a quarterly sales report where each product’s contribution is shown as a stacked column. | Build a financial dashboard Excel file that visualizes cumulative revenue across quarters. | Export sales data with a pre‑formatted stacked column chart for stakeholder presentations.
// AI Prompts: Show how to add data labels to each segment of the stacked column chart created with Aspose.Cells. | Convert the stacked column chart to a 100 % stacked column chart while preserving the data range. | Demonstrate applying a custom color palette to each series in the ColumnStacked chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace StackedColumnChartExample
{
    // This C# sample builds a new workbook, fills it with quarterly sales data for three products, adds a ColumnStacked chart (quarters on the X‑axis, sales values as series), sets a title, and saves the file as StackedColumnChart.xlsx using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: cumulative sales for three products across four quarters
            // Header row
            sheet.Cells["A1"].PutValue("Quarter");
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["C1"].PutValue("Product B");
            sheet.Cells["D1"].PutValue("Product C");

            // Data rows
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["C2"].PutValue(80);
            sheet.Cells["D2"].PutValue(50);

            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["C3"].PutValue(110);
            sheet.Cells["D3"].PutValue(70);

            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["C4"].PutValue(130);
            sheet.Cells["D4"].PutValue(90);

            sheet.Cells["A5"].PutValue("Q4");
            sheet.Cells["B5"].PutValue(200);
            sheet.Cells["C5"].PutValue(150);
            sheet.Cells["D5"].PutValue(110);

            // Add a stacked column chart (cumulative view)
            // Parameters: chart type, top row, left column, bottom row, right column
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 7, 0, 25, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (sales values)
            // The first parameter is the range for the series values; true indicates that the range is by column
            chart.NSeries.Add("B2:D5", true);
            // Set the category (X‑axis) data – the quarters
            chart.NSeries.CategoryData = "A2:A5";

            // Optional: give the chart a title
            chart.Title.Text = "Cumulative Quarterly Sales (Stacked Column)";

            // Save the workbook to a file
            workbook.Save("StackedColumnChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
