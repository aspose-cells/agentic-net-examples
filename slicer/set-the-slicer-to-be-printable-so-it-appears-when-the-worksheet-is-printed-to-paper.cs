using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

class SetSlicerPrintable
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].Value = "Category";
        cells["A2"].Value = "A";
        cells["A3"].Value = "B";
        cells["A4"].Value = "A";

        cells["B1"].Value = "Value";
        cells["B2"].Value = 10;
        cells["B3"].Value = 20;
        cells["B4"].Value = 30;

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "Pivot1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the pivot field "Category"
        int slicerIndex = sheet.Slicers.Add(pivot, "E1", "Category");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Ensure the slicer is printable by setting the underlying shape's IsPrintable property
        slicer.Shape.IsPrintable = true;

        // Save the workbook
        workbook.Save("SlicerPrintable.xlsx");
    }
}