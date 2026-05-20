using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class CloneSlicerExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -----------------------------------------------------------------
        // Source worksheet: contains data, pivot table and the original slicer
        // -----------------------------------------------------------------
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Populate sample data
        sourceSheet.Cells["A1"].PutValue("Fruit");
        sourceSheet.Cells["B1"].PutValue("Sales");
        sourceSheet.Cells["A2"].PutValue("Apple");
        sourceSheet.Cells["B2"].PutValue(100);
        sourceSheet.Cells["A3"].PutValue("Orange");
        sourceSheet.Cells["B3"].PutValue(150);
        sourceSheet.Cells["A4"].PutValue("Banana");
        sourceSheet.Cells["B4"].PutValue(200);

        // Create a pivot table based on the data
        int pivotIndex = sourceSheet.PivotTables.Add("A1:B4", "D2", "Pivot1");
        PivotTable pivot = sourceSheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
        pivot.RefreshData();
        pivot.CalculateData();

        // Add the original slicer on the source sheet
        int originalSlicerIndex = sourceSheet.Slicers.Add(pivot, "F2", "Fruit");
        Slicer originalSlicer = sourceSheet.Slicers[originalSlicerIndex];
        originalSlicer.Caption = "Fruit Filter";
        originalSlicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        originalSlicer.NumberOfColumns = 1;
        originalSlicer.WidthPixel = 150;
        originalSlicer.HeightPixel = 200;

        // ---------------------------------------------------------------
        // Destination worksheet: will host the cloned slicer for parallel use
        // ---------------------------------------------------------------
        Worksheet destSheet = workbook.Worksheets.Add("Clone");

        // Add a slicer on the destination sheet using the same pivot table and field
        int clonedSlicerIndex = destSheet.Slicers.Add(pivot, "F2", "Fruit");
        Slicer clonedSlicer = destSheet.Slicers[clonedSlicerIndex];

        // Copy visual and behavioral properties from the original slicer
        clonedSlicer.Caption = originalSlicer.Caption;
        clonedSlicer.StyleType = originalSlicer.StyleType;
        clonedSlicer.NumberOfColumns = originalSlicer.NumberOfColumns;
        clonedSlicer.WidthPixel = originalSlicer.WidthPixel;
        clonedSlicer.HeightPixel = originalSlicer.HeightPixel;
        clonedSlicer.LockedPosition = originalSlicer.LockedPosition;
        clonedSlicer.ShowCaption = originalSlicer.ShowCaption;

        // Ensure the cloned slicer is synchronized with its pivot table
        clonedSlicer.Refresh();

        // Save the workbook
        workbook.Save("ClonedSlicerDemo.xlsx");
    }
}