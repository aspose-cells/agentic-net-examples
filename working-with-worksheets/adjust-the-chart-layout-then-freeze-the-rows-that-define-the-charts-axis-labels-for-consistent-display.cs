using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLayoutAndFreeze
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            // Category labels (axis labels) in column A
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["A5"].PutValue("D");

            // Corresponding values in column B
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);
            worksheet.Cells["B5"].PutValue(40);

            // Add a column chart positioned below the data range
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the chart data range (including headers)
            chart.SetChartDataRange("A1:B5", true);

            // Adjust chart layout
            // 1. Ensure the chart moves and resizes with cells
            chart.Placement = PlacementType.MoveAndSize;

            // 2. Enable automatic resizing with the window (optional)
            chart.SizeWithWindow = true;

            // 3. Reset plot area to automatic positioning
            chart.PlotArea.SetPositionAuto();

            // 4. Recalculate the chart to apply layout changes
            chart.Calculate();

            // Freeze the rows that contain the axis (category) labels
            // Axis labels are in rows 1‑5, so freeze the first 5 rows.
            // Freeze at cell "A6" (first row after the labels) with 5 frozen rows.
            worksheet.FreezePanes("A6", 5, 0);

            // Save the workbook
            workbook.Save("ChartLayoutAndFreeze.xlsx");
        }
    }
}