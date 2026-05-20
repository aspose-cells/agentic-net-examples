using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

class MoveSlicerAndAlignChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ---------- Sample data ----------
        sheet.Cells["A1"].Value = "Fruit";
        sheet.Cells["A2"].Value = "Apple";
        sheet.Cells["A3"].Value = "Orange";
        sheet.Cells["A4"].Value = "Banana";

        sheet.Cells["B1"].Value = "Sales";
        sheet.Cells["B2"].Value = 120;
        sheet.Cells["B3"].Value = 150;
        sheet.Cells["B4"].Value = 200;

        // ---------- Create a pivot table ----------
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "E1", "FruitPivot");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // ---------- Add a chart (for demonstration) ----------
        int chartIdx = sheet.Charts.Add(ChartType.Column, 10, 0, 20, 5);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // ---------- Add a slicer linked to the pivot ----------
        int slicerIdx = sheet.Slicers.Add(pivot, "A6", "Fruit");
        Slicer slicer = sheet.Slicers[slicerIdx];

        // ---------- Move slicer to cell D5 ----------
        // D5 => row index 4 (zero‑based), column index 3 (zero‑based)
        // Use the Shape object of the slicer to set its upper‑left cell position
        slicer.Shape.UpperLeftRow = 4;      // Row 5 in Excel
        slicer.Shape.UpperLeftColumn = 3;   // Column D in Excel

        // ---------- Align the existing chart with the slicer ----------
        // Keep the chart size unchanged, only move its top‑left corner to D5
        // Retrieve current size of the chart
        int currentBottomRow = chart.ChartObject.LowerRightRow;
        int currentRightColumn = chart.ChartObject.LowerRightColumn;

        // Move the chart so its upper‑left corner matches the slicer's position
        chart.Move(4, 3, currentBottomRow, currentRightColumn);

        // ---------- Save the workbook ----------
        workbook.Save("SlicerAndChartAligned.xlsx");
    }
}