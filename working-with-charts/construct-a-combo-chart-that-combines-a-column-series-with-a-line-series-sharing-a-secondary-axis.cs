// Title: Create a Combo Chart (Column + Line) with a Secondary Axis using Aspose.Cells for .NET
// Description: C# example that builds a workbook, adds category data, creates a column chart, converts the second series to a line, plots it on a secondary Y‑axis, sets an optional axis title, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | combo chart | column series | line series | secondary axis | plot on second axis | chart series type | Excel chart automation | chart customization
// Common Searches: Aspose.Cells combo chart column and line | how to add a secondary axis to a chart in Aspose.Cells | change series type to line in Aspose.Cells chart | C# create combo chart with two axes using Aspose.Cells | plot line series on secondary Y axis Aspose.Cells
// Developer Intent: Generate a combo chart that mixes column and line series, with the line series displayed on a secondary Y‑axis.
// Use Cases: Display monthly sales (columns) alongside profit margin percentage (line) on separate scales. | Compare production volume (columns) with defect rate (line) in a manufacturing KPI report. | Show website traffic (columns) together with conversion rate (line) in a marketing dashboard.
// AI Prompts: Write C# code with Aspose.Cells to create a combo chart that combines a column series and a line series on a secondary axis, including axis titles. | Explain how to switch a specific chart series to a line type and assign it to the secondary Y‑axis in Aspose.Cells. | Provide step‑by‑step instructions to customize the secondary value axis title for a combo chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ComboChartExample
{
    // C# example that builds a workbook, adds category data, creates a column chart, converts the second series to a line, plots it on a secondary Y‑axis, sets an optional axis title, and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A – Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            // Column B – Values for the column series
            sheet.Cells["B1"].PutValue("ColumnSeries");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Column C – Values for the line series
            sheet.Cells["C1"].PutValue("LineSeries");
            sheet.Cells["C2"].PutValue(100);
            sheet.Cells["C3"].PutValue(150);
            sheet.Cells["C4"].PutValue(200);

            // Add a combo chart (initially a Column chart)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Add the first series (column) and the second series (line)
            chart.NSeries.Add("B2:B4", true); // first series
            chart.NSeries.Add("C2:C4", true); // second series

            // Set the category (X) axis data
            chart.NSeries.CategoryData = "A2:A4";

            // Change the second series to a line type
            chart.NSeries[1].Type = ChartType.Line;

            // Plot the line series on the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Optional: customize the secondary axis title
            chart.SecondValueAxis.Title.Text = "Secondary Axis";

            // Save the workbook
            workbook.Save("ComboChart_ColumnLine_SecondaryAxis.xlsx");
        }
    }
}
