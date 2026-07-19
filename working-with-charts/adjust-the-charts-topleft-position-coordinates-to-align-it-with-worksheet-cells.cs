// Title: C# – Align an Aspose.Cells Chart to Worksheet Cells with Chart.Move and PlacementType
// Description: This example creates a workbook, fills A1:B4 with sample data, adds a column chart, and uses Chart.Move(1, 2, 11, 7) to position the chart’s top‑left corner at cell C2 and bottom‑right corner at cell H12. The chart’s Placement is set to MoveAndSize so it follows cell changes, and the file is saved as ChartAlignedWithCells.xlsx.
// Keywords: Aspose.Cells | Chart.Move | PlacementType | C# chart positioning | align chart to cells | move and size chart | Aspose.Cells example | chart placement .NET | Excel chart coordinates | Aspose.Cells API
// Common Searches: Aspose.Cells chart Move method example | how to position a chart at cell C2 in .NET | set chart placement to MoveAndSize Aspose.Cells | align chart corners to specific cells using Aspose.Cells | C# Aspose.Cells chart positioning tutorial
// Developer Intent: The developer needs to place a chart so its top‑left and bottom‑right corners match specific worksheet cells and have the chart automatically move and resize with those cells.
// Use Cases: Embed a column chart over a data table and lock its location to cells C2:H12, preserving layout when rows or columns are added. | Build a dynamic dashboard where charts expand or shrink as the underlying data range grows, using PlacementType.MoveAndSize. | Generate printable Excel reports with charts precisely aligned to header cells for consistent visual formatting.
// AI Prompts: Show C# code that uses Chart.Move to align a chart with cells A5:D15 and sets Placement to MoveAndSize in Aspose.Cells. | Provide a snippet that repositions a chart after inserting rows so it stays aligned with the range C2:H12. | Explain the differences between PlacementType.MoveAndSize, PlacementType.Move, and PlacementType.FreeFloating for Aspose.Cells charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // Required for PlacementType enum

namespace AsposeCellsChartPositionDemo
{
    // This example creates a workbook, fills A1:B4 with sample data, adds a column chart, and uses Chart.Move(1, 2, 11, 7) to position the chart’s top‑left corner at cell C2 and bottom‑right corner at cell H12. The chart’s Placement is set to MoveAndSize so it follows cell changes, and the file is saved as ChartAlignedWithCells.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["B4"].PutValue(25);

                // Add a column chart. Initial position is arbitrary.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Align the chart's top‑left corner to cell C2 (row 1, column 2)
                // and bottom‑right corner to cell H12 (row 11, column 7)
                // Chart.Move expects (topRow, leftColumn, bottomRow, rightColumn)
                chart.Move(1, 2, 11, 7);

                // Optional: make the chart move and size with cells
                chart.Placement = PlacementType.MoveAndSize;

                // Save the workbook (lifecycle rule: save)
                workbook.Save("ChartAlignedWithCells.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
