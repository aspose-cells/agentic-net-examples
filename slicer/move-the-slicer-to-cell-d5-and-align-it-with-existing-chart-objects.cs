// Title: Reposition a slicer to cell D5 and align it with a chart using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that sets a slicer's UpperLeftColumn and UpperLeftRow so its top‑left corner sits on cell D5. | Show how to lock a slicer's position after moving it and ensure it aligns with an existing chart on the same worksheet using Aspose.Cells. | Demonstrate moving a slicer, adjusting its placement type, and saving the workbook with the updated layout in C#.
// Common Searches: Aspose.Cells C# how to place a slicer at cell D5 | move slicer to specific cell and align with chart Aspose.Cells | set slicer UpperLeftColumn UpperLeftRow Aspose.Cells example | lock slicer position after moving in Aspose.Cells for .NET | align slicer and chart placement using Aspose.Cells C#
// Tags: Aspose.Cells C# move slicer location | Aspose.Cells set slicer UpperLeftColumn UpperLeftRow | Aspose.Cells align slicer with chart | Aspose.Cells lock slicer position | Aspose.Cells pivot slicer positioning

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, builds a pivot table and a column chart, inserts a slicer for the pivot, moves the slicer so its upper‑left corner aligns with cell D5, locks the slicer's position, and saves the file as MovedSlicer.xlsx.
class MoveSlicerExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Fruit";
        sheet.Cells["A2"].Value = "Apple";
        sheet.Cells["A3"].Value = "Orange";
        sheet.Cells["A4"].Value = "Banana";
        sheet.Cells["B1"].Value = "Sales";
        sheet.Cells["B2"].Value = 100;
        sheet.Cells["B3"].Value = 150;
        sheet.Cells["B4"].Value = 200;

        // Add a pivot table based on the data
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "E1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Add a chart that uses the same data range
        int chartIdx = sheet.Charts.Add(ChartType.Column, 2, 5, 12, 10);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Ensure the chart moves and sizes with cells (optional alignment step)
        chart.Placement = PlacementType.MoveAndSize;

        // Add a slicer for the pivot table (initially placed at A6)
        int slicerIdx = sheet.Slicers.Add(pivot, "A6", "Fruit");
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Move the slicer so its upper‑left corner aligns with cell D5
        // D5 corresponds to column index 3 (A=0) and row index 4 (1‑based row 5)
        slicer.Shape.UpperLeftColumn = 3; // Column D
        slicer.Shape.UpperLeftRow = 4;    // Row 5

        // Optionally lock the slicer position so users cannot move it unintentionally
        slicer.LockedPosition = true;

        // Save the workbook
        workbook.Save("MovedSlicer.xlsx");
    }
}
