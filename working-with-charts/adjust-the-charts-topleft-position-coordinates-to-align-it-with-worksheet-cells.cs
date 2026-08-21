// Title: C# – Align an Aspose.Cells chart with specific worksheet cells using Chart.Move
// Description: This example creates a workbook, fills cells A1:B4 with sample data, adds a column chart, and then uses the Chart.Move method to place the chart’s top‑left corner at cell C2 and its bottom‑right corner at cell H12. The chart’s Placement is set to MoveAndSize so it moves and resizes with the cells, and the file is saved as AdjustedChart.xlsx.
// Keywords: Aspose.Cells chart positioning | Chart.Move C# | PlacementType.MoveAndSize | align chart with cells Aspose.Cells | .NET chart placement | adjust chart top left cell | resize chart with worksheet rows columns
// Common Searches: Aspose.Cells move chart to cell C2 | Chart.Move top left and bottom right coordinates | Set chart to move and resize with cells Aspose.Cells | Align Aspose.Cells chart boundaries with specific cells | C# chart placement based on row and column indices
// Developer Intent: Place and size a chart so its corners match designated worksheet cells and ensure the chart follows any row or column changes.
// Use Cases: Overlay a column chart precisely over a data table for a clean report layout. | Build a dynamic dashboard where charts stay anchored to defined cell ranges despite insertions or deletions. | Generate financial statements with charts that automatically adjust to the surrounding cell block.
// AI Prompts: Write C# code with Aspose.Cells to move a chart to the range D5:G15 and set its placement to MoveAndSize. | Explain how each parameter of Chart.Move maps to worksheet rows and columns in Aspose.Cells. | Show how to programmatically recalculate a chart’s position when the underlying data range expands.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This example creates a workbook, fills cells A1:B4 with sample data, adds a column chart, and then uses the Chart.Move method to place the chart’s top‑left corner at cell C2 and its bottom‑right corner at cell H12. The chart’s Placement is set to MoveAndSize so it moves and resizes with the cells, and the file is saved as AdjustedChart.xlsx.
class AdjustChartPosition
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a chart (initial position does not matter)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Move the chart so its top‑left corner aligns with cell C2 (row 1, column 2)
        // and its bottom‑right corner aligns with cell H12 (row 11, column 7)
        chart.Move(topRow: 1, leftColumn: 2, bottomRow: 11, rightColumn: 7);

        // Ensure the chart moves and resizes together with the cells
        chart.Placement = PlacementType.MoveAndSize;

        // Save the workbook with the adjusted chart position
        workbook.Save("AdjustedChart.xlsx");
    }
}
