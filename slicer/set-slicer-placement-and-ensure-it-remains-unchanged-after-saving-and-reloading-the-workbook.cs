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
            // ---------- Create a new workbook and add data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for a pivot table
            sheet.Cells["A1"].Value = "Fruit";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B1"].Value = "Sales";
            sheet.Cells["B2"].Value = 100;
            sheet.Cells["B3"].Value = 150;
            sheet.Cells["B4"].Value = 200;

            // ---------- Create a pivot table ----------
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // ---------- Add a slicer linked to the pivot table ----------
            // Destination cell for the slicer is A6, base field index is 0 (Fruit)
            int slicerIdx = sheet.Slicers.Add(pivot, "A6", 0);
            Slicer slicer = sheet.Slicers[slicerIdx];

            // ---------- Set the slicer placement ----------
            // Using the obsolete Slicer.Placement property (as per available rule)
            slicer.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            string filePath = "SlicerPlacementDemo.xlsx";
            workbook.Save(filePath);

            // ---------- Reload the workbook and verify placement ----------
            Workbook loadedWb = new Workbook(filePath);
            Worksheet loadedSheet = loadedWb.Worksheets[0];
            Slicer loadedSlicer = loadedSheet.Slicers[0];

            // Check if the placement persisted
            PlacementType placement = loadedSlicer.Placement;
            Console.WriteLine("Slicer placement after reload: " + placement);
            // Expected output: MoveAndSize
        }
    }
}