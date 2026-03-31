using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace SlicerSettingsDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputData.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(100);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(200);
            cells["A4"].PutValue("A");
            cells["B4"].PutValue(150);
            cells["A5"].PutValue("C");
            cells["B5"].PutValue(120);

            // Create a table from the data range
            int tableIdx = sheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = sheet.ListObjects[tableIdx];

            // Add a pivot table based on the table's data range
            PivotTableCollection pivots = sheet.PivotTables;
            // Use the table's data range as a string (e.g., "A1:B5")
            string sourceRange = $"{cells[0, 0].Name}:{cells[4, 1].Name}";
            int pivotIdx = pivots.Add(sourceRange, "D2", "PivotDemo");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table's "Category" field
            SlicerCollection slicers = sheet.Slicers;
            int slicerIdx = slicers.Add(pivot, "G2", "Category");
            Slicer slicer = slicers[slicerIdx];

            // Adjust slicer settings
            slicer.ShowCaption = false;
            slicer.Caption = "Category Filter";
            slicer.ShowAllItems = false;
            slicer.ShowMissing = true;
            slicer.NumberOfColumns = 2;
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.LockedPosition = false;
            // Use Shape property to lock aspect ratio
            slicer.Shape.IsLockAspectRatio = true;

            slicer.Refresh();

            // Save the modified workbook
            workbook.Save("SlicerSettingsDemo.xlsx");
        }
    }
}