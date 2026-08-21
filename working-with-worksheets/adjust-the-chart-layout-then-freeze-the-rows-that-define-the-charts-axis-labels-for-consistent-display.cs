// Title: C# Aspose.Cells: Relocate a Chart with MoveAndSize Placement and Freeze Axis‑Label Rows
// Description: This example creates a workbook, fills columns A and B with sample categories and values, adds a column chart, sets its placement to MoveAndSize so it follows cell changes, recalculates the chart, moves it to rows 12‑22 and columns 2‑10, then freezes the first six rows (the axis‑label rows) using FreezePanes at cell A7, and finally saves the file as ChartLayoutAndFreeze.xlsx.
// Keywords: Aspose.Cells chart placement | MoveAndSize | chart.Move C# | FreezePanes Aspose.Cells | freeze worksheet rows | axis label rows | C# Aspose.Cells example | adjust chart layout programmatically | worksheet freeze panes .NET | column chart Aspose.Cells
// Common Searches: Aspose.Cells move chart with cells .NET | Set chart placement to MoveAndSize in C# | How to relocate a chart programmatically using Aspose.Cells | Freeze specific rows in an Aspose.Cells worksheet | Freeze panes below chart axis labels Aspose.Cells | C# example for chart.Calculate and chart.Move | Aspose.Cells FreezePanes at a given cell
// Developer Intent: The developer wants to move a chart so it follows cell modifications and then keep the rows used for axis labels fixed while scrolling.
// Use Cases: Align a chart with dynamic data ranges by using MoveAndSize placement. | Shift a chart to a different area of the sheet after inserting new rows or columns. | Keep header rows that serve as chart axis labels frozen for better readability in large reports. | Ensure consistent chart appearance across different screen sizes by recalculating its position before moving. | Generate Excel reports where both the chart and its axis labels remain static during navigation.
// AI Prompts: Generate C# code that sets an Aspose.Cells chart’s Placement to MoveAndSize, calls Calculate, and moves it to a specific cell range. | Show how to use Worksheet.FreezePanes to lock the first N rows while leaving columns unfrozen in Aspose.Cells for .NET. | Explain the steps to adjust a chart’s layout and then freeze the rows containing its axis labels using Aspose.Cells. | Provide a complete Aspose.Cells example that creates data, adds a column chart, relocates it, and freezes the axis‑label rows.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLayoutAndFreeze
{
    // This example creates a workbook, fills columns A and B with sample categories and values, adds a column chart, sets its placement to MoveAndSize so it follows cell changes, recalculates the chart, moves it to rows 12‑22 and columns 2‑10, then freezes the first six rows (the axis‑label rows) using FreezePanes at cell A7, and finally saves the file as ChartLayoutAndFreeze.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (categories in column A, values in column B)
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 6; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                worksheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a column chart and set its data range
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 8, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.SetChartDataRange("A1:B6", true);

            // Adjust chart layout:
            // 1. Ensure the chart moves and resizes with cells
            chart.Placement = PlacementType.MoveAndSize;
            // 2. Recalculate positions (use Calculate method)
            chart.Calculate();
            // 3. Move the chart to a new location (rows 12‑22, columns 2‑10)
            chart.Move(12, 2, 22, 10);

            // Freeze the rows that contain the axis labels (rows 1‑6)
            // Freeze panes just below row 6 (cell A7) with 6 frozen rows and 0 frozen columns
            worksheet.FreezePanes("A7", 6, 0);

            // Save the workbook
            workbook.Save("ChartLayoutAndFreeze.xlsx");
        }
    }
}
