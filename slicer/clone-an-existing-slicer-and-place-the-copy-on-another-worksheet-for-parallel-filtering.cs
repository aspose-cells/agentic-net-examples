using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerCloneDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a workbook with sample data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "DataSheet";

            // Populate sample data
            sheet1.Cells["A1"].PutValue("Fruit");
            sheet1.Cells["B1"].PutValue("Quantity");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["B2"].PutValue(10);
            sheet1.Cells["A3"].PutValue("Orange");
            sheet1.Cells["B3"].PutValue(20);
            sheet1.Cells["A4"].PutValue("Banana");
            sheet1.Cells["B4"].PutValue(15);

            // ---------- Create a pivot table ----------
            int pivotIdx = sheet1.PivotTables.Add("A1:B4", "D1", "FruitPivot");
            PivotTable pivot = sheet1.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // ---------- Add an original slicer on the first worksheet ----------
            int slicerIdx = sheet1.Slicers.Add(pivot, "E2", "Fruit");
            Slicer originalSlicer = sheet1.Slicers[slicerIdx];

            // Set some visual properties on the original slicer
            originalSlicer.Caption = "Fruit Filter";
            originalSlicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            originalSlicer.NumberOfColumns = 2;
            originalSlicer.WidthPixel = 200;
            originalSlicer.HeightPixel = 150;
            originalSlicer.LockedPosition = false;

            // ---------- Add a second worksheet where the slicer will be cloned ----------
            Worksheet sheet2 = workbook.Worksheets.Add("CloneSheet");

            // ---------- Clone the slicer ----------
            // Add a new slicer on the second worksheet using the same pivot table and base field
            int clonedSlicerIdx = sheet2.Slicers.Add(pivot, "E2", "Fruit");
            Slicer clonedSlicer = sheet2.Slicers[clonedSlicerIdx];

            // Copy visual properties from the original slicer to the cloned slicer
            clonedSlicer.Caption = originalSlicer.Caption;
            clonedSlicer.StyleType = originalSlicer.StyleType;
            clonedSlicer.NumberOfColumns = originalSlicer.NumberOfColumns;
            clonedSlicer.WidthPixel = originalSlicer.WidthPixel;
            clonedSlicer.HeightPixel = originalSlicer.HeightPixel;
            clonedSlicer.LockedPosition = originalSlicer.LockedPosition;

            // ---------- Save the workbook ----------
            workbook.Save("SlicerCloneDemo.xlsx");
        }
    }
}