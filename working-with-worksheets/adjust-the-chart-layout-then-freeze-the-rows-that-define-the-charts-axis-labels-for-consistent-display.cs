// Title: C# Aspose.Cells – Set Chart Placement to MoveAndSize and Freeze Header Rows with Axis Labels
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, configure the chart to move and resize with cells (Placement = MoveAndSize), recalculate its layout, reposition it, and then freeze the top rows that contain the chart's axis labels using FreezePanes, before saving the file.
// Keywords: Aspose.Cells chart placement | MoveAndSize placement | freeze top rows Aspose.Cells | FreezePanes axis labels | C# Excel chart layout | recalculate chart layout Aspose | programmatic chart reposition | Excel freeze panes C# | Aspose.Cells worksheet formatting
// Common Searches: Aspose.Cells set chart to move and size with cells | How to freeze rows that contain chart axis labels in .NET | C# code to reposition an Excel chart using Aspose.Cells | FreezePanes example for header rows in Aspose.Cells | Adjust chart layout and freeze panes Aspose.Cells
// Developer Intent: Configure a chart to stay aligned with its cells and keep the axis‑label rows visible by freezing them in the worksheet.
// Use Cases: Create a column chart, set its Placement to MoveAndSize, call Calculate, and move it to a new cell range. | Freeze the first four rows (the category and value headers that serve as axis labels) while leaving columns scrollable. | Maintain consistent display of chart axis labels during vertical scrolling in generated Excel files.
// AI Prompts: Generate C# code with Aspose.Cells that sets a chart's Placement to MoveAndSize, recalculates its layout, moves the chart to a specific range, and freezes the top N rows. | Show an example of using FreezePanes in Aspose.Cells to lock header rows without freezing columns. | Explain how to combine chart layout adjustments and FreezePanes to keep axis labels visible when scrolling in an Excel worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLayoutAndFreeze
{
    // Demonstrates how to create a workbook, add sample data, insert a column chart, configure the chart to move and resize with cells (Placement = MoveAndSize), recalculate its layout, reposition it, and then freeze the top rows that contain the chart's axis labels using FreezePanes, before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (category labels in column A, values in column B)
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart covering rows 5‑20 and columns 0‑8
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Adjust chart layout:
            // 1. Ensure the chart moves and resizes with cells
            chart.Placement = PlacementType.MoveAndSize;
            // 2. Recalculate automatic positions (plot area, axes, etc.)
            chart.Calculate();
            // 3. Optionally move the chart to a different area (example)
            chart.Move(10, 2, 22, 10);

            // Freeze the rows that contain the axis labels (rows 1‑4)
            // Freeze at the first unfrozen row (row index 5, column index 0)
            // This freezes the top 4 rows while leaving columns unfrozen
            worksheet.FreezePanes(5, 0, 4, 0);

            // Save the workbook
            workbook.Save("ChartLayoutAndFreeze.xlsx");
        }
    }
}
