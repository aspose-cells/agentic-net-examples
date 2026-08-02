// Title: C# – Aspose.Cells: Relocate PivotTable slicer to D5 and sync its size with a chart
// Description: Demonstrates creating a workbook, adding sample data, building a PivotTable, inserting a column chart, creating a slicer for the "Category" field, then using UpperLeftRow/UpperLeftColumn to place the slicer at D5 and copying the chart's Height and Width to the slicer's Shape before saving.
// Keywords: Aspose.Cells .NET | C# slicer positioning | UpperLeftRow UpperLeftColumn | slicer resize to chart | PivotTable slicer alignment | Excel automation Aspose | move slicer programmatically | chart dimensions Aspose.Cells | Excel dashboard layout
// Common Searches: Aspose.Cells move slicer to specific cell C# | How to align slicer size with a chart using Aspose.Cells | Set slicer UpperLeftRow and UpperLeftColumn in .NET | Resize slicer to match chart dimensions programmatically | PivotTable slicer placement with Aspose.Cells
// Developer Intent: Programmatically place a PivotTable slicer at cell D5 and make its dimensions identical to an existing chart in a .NET workbook.
// Use Cases: Design compact dashboards where slicers sit next to charts for a clean UI. | Automate report generation that requires slicer alignment across multiple worksheets. | Adjust slicer location dynamically when new data ranges or visual elements are added.
// AI Prompts: Write C# code with Aspose.Cells to position a slicer at D5 and set its Height and Width equal to a given chart. | Explain how to retrieve a chart's Height and Width properties and apply them to a slicer's Shape in Aspose.Cells for .NET. | Provide best‑practice error handling for slicer placement and resizing in automated Excel workbook creation.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, adding sample data, building a PivotTable, inserting a column chart, creating a slicer for the "Category" field, then using UpperLeftRow/UpperLeftColumn to place the slicer at D5 and copying the chart's Height and Width to the slicer's Shape before saving.
class MoveSlicerAndAlignWithChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------- Sample data for PivotTable --------------------
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["A4"].Value = "Vegetable";
            sheet.Cells["B1"].Value = "Sales";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["B3"].Value = 150;
            sheet.Cells["B4"].Value = 200;

            // -------------------- Create PivotTable --------------------
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "E1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------- Add a chart linked to the PivotTable --------------------
            int chartIdx = sheet.Charts.Add(ChartType.Column, 2, 0, 12, 5);
            Chart chart = sheet.Charts[chartIdx];
            // Use the same data range as the pivot for demonstration
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // -------------------- Add a slicer for the PivotTable --------------------
            int slicerIdx = sheet.Slicers.Add(pivot, "G2", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // -------------------- Move slicer to cell D5 --------------------
            int targetRow = sheet.Cells["D5"].Row;          // zero‑based row index
            int targetColumn = sheet.Cells["D5"].Column;    // zero‑based column index
            slicer.Shape.UpperLeftRow = targetRow;
            slicer.Shape.UpperLeftColumn = targetColumn;

            // -------------------- Align slicer size with the existing chart --------------------
            // Directly use the ChartObject properties without declaring a separate variable
            slicer.Shape.Height = chart.ChartObject.Height;
            slicer.Shape.Width = chart.ChartObject.Width;

            // -------------------- Save the workbook --------------------
            workbook.Save("SlicerMovedAndAligned.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
