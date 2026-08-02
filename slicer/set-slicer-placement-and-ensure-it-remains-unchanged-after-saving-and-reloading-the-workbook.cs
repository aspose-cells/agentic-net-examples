using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace SlicerPlacementDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create workbook and add data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for pivot table
            sheet.Cells["A1"].Value = "Fruit";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B1"].Value = "Sales";
            sheet.Cells["B2"].Value = 100;
            sheet.Cells["B3"].Value = 150;
            sheet.Cells["B4"].Value = 200;

            // ---------- Create pivot table ----------
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D3", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // ---------- Add slicer linked to the pivot table ----------
            // Destination cell for slicer upper‑left corner is "F3"
            int slicerIdx = sheet.Slicers.Add(pivot, "F3", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // ---------- Set slicer placement ----------
            // Use the (obsolete) Placement property as required
            slicer.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            string filePath = "SlicerPlacementDemo.xlsx";
            workbook.Save(filePath);

            // ---------- Reload workbook and verify placement ----------
            Workbook loadedWb = new Workbook(filePath);
            Worksheet loadedSheet = loadedWb.Worksheets[0];
            Slicer loadedSlicer = loadedSheet.Slicers[0];

            // Output the placement value to confirm it persisted
            Console.WriteLine("Slicer placement after reload: " + loadedSlicer.Placement);
        }
    }
}