using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].PutValue("Fruit");
            cells["B1"].PutValue("Quantity");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(50);
            cells["A3"].PutValue("Orange");
            cells["B3"].PutValue(30);
            cells["A4"].PutValue("Banana");
            cells["B4"].PutValue(20);
            cells["A5"].PutValue("Apple");
            cells["B5"].PutValue(40);
            cells["A6"].PutValue("Orange");
            cells["B6"].PutValue(25);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table
            // Destination cell for the slicer is E3, base field name is "Fruit"
            int slicerIndex = sheet.Slicers.Add(pivot, "E3", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Modify slicer properties
            slicer.Caption = "Fruit Slicer";
            slicer.LockedPosition = true;                     // Prevent moving/resizing via UI
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

            // Refresh the slicer (also refreshes the underlying pivot table)
            slicer.Refresh();

            // Export the workbook to an XLSX file
            string filePath = "ModifiedSlicer.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Reload the saved workbook from the file
            Workbook reloadedWorkbook = new Workbook(filePath);

            // Access the slicer in the reloaded workbook to verify it exists
            Worksheet reloadedSheet = reloadedWorkbook.Worksheets[0];
            Slicer reloadedSlicer = reloadedSheet.Slicers[0];
            Console.WriteLine("Reloaded slicer caption: " + reloadedSlicer.Caption);
        }
    }
}