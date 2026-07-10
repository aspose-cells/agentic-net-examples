using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerWithPivotExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Quantity";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 10;
            cells["A3"].Value = "Orange";
            cells["B3"].Value = 5;
            cells["A4"].Value = "Banana";
            cells["B4"].Value = 8;

            // Add a pivot table based on the data range A1:B4, place it at C3
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "C3", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Configure the pivot: Fruit as row field, Quantity as data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table.
            // The slicer will be placed with its upper‑left corner at cell E2
            // and will filter by the "Fruit" field.
            int slicerIdx = sheet.Slicers.Add(pivot, "E2", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Optional: customize slicer appearance
            slicer.Caption = "Fruit Filter";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.NumberOfColumns = 1;
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 120;

            // Refresh the slicer to ensure it reflects the current pivot data
            slicer.Refresh();

            // Save the workbook
            workbook.Save("SlicerLinkedToPivot.xlsx");
        }
    }
}