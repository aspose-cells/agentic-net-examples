// Title: Aspose.Cells C# – Configure Chart Axes and Freeze Rows Containing Axis Labels
// Description: This example creates a workbook, adds sample sales data, inserts a clustered column chart, customizes the X and Y axis titles, tick marks, and value range, calculates the chart to generate labels, and then freezes the top rows that hold the axis titles and category labels before saving the file.
// Keywords: Aspose.Cells | C# chart axis | Aspose.Cells chart axis title | FreezePanes C# | freeze rows Aspose.Cells | set chart min max value | custom tick marks Aspose.Cells | .NET chart example | worksheet freeze panes | chart axis labels | column chart Aspose.Cells | programmatic chart formatting
// Common Searches: Aspose.Cells set X axis title C# | Aspose.Cells set Y axis min max values | FreezePanes top rows after chart Aspose.Cells | How to keep chart axis labels visible while scrolling | C# example for freezing rows with chart labels | Aspose.Cells calculate chart before freezing panes
// Developer Intent: I want to programmatically define chart axis properties and keep the rows that contain those axis labels fixed while the user scrolls.
// Use Cases: Generate a sales column chart with custom axis titles, tick marks, and a fixed value range, then freeze the header rows so the axis context stays in view. | Create a financial worksheet that includes a chart and requires the axis label rows to remain pinned during vertical scrolling for better readability. | Build an automated dashboard that formats chart axes via code and applies FreezePanes to preserve axis information for end‑users.
// AI Prompts: Write C# code using Aspose.Cells to set category and value axis titles, adjust tick marks, define min/max values, calculate the chart, and freeze the first three rows of the worksheet. | Provide an Aspose.Cells example that configures a column chart’s value axis with a minimum of 0, maximum of 200, major unit of 50, and then applies FreezePanes to keep axis labels visible. | Explain step‑by‑step how to calculate a chart after setting axis properties to ensure labels are generated before calling FreezePanes in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAxisFreezeDemo
{
    // This example creates a workbook, adds sample sales data, inserts a clustered column chart, customizes the X and Y axis titles, tick marks, and value range, calculates the chart to generate labels, and then freezes the top rows that hold the axis titles and category labels before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            sheet.Cells["C1"].PutValue("Product B");
            sheet.Cells["C2"].PutValue(80);
            sheet.Cells["C3"].PutValue(130);
            sheet.Cells["C4"].PutValue(170);

            // Add a clustered column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the chart
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure Category (X) axis
            Axis catAxis = chart.CategoryAxis;
            catAxis.Title.Text = "Quarter";
            catAxis.Title.IsVisible = true;
            catAxis.MajorTickMark = TickMarkType.Outside;
            catAxis.MinorTickMark = TickMarkType.Inside;
            catAxis.TickLabelPosition = TickLabelPositionType.NextToAxis;

            // Configure Value (Y) axis
            Axis valAxis = chart.ValueAxis;
            valAxis.Title.Text = "Sales (Units)";
            valAxis.Title.IsVisible = true;
            valAxis.IsAutomaticMinValue = false;
            valAxis.MinValue = 0;
            valAxis.IsAutomaticMaxValue = false;
            valAxis.MaxValue = 200;
            valAxis.MajorUnit = 50;
            valAxis.MajorTickMark = TickMarkType.Outside;
            valAxis.MinorTickMark = TickMarkType.Inside;

            // Optional: calculate the chart to ensure axis labels are generated
            chart.Calculate();

            // Freeze the top three rows (which contain the axis titles and category labels)
            // FreezePanes(string cellName, int freezedRows, int freezedColumns)
            sheet.FreezePanes("A4", 3, 0);

            // Save the workbook
            workbook.Save("ChartAxisFreezeDemo.xlsx");
        }
    }
}
