using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerManagementDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("Input.xlsx");

            // Access the first worksheet (adjust index/name as needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // Ensure there is a PivotTable to which slicers can be attached.
            // If the worksheet already contains a PivotTable, you can skip this block.
            // -------------------------------------------------
            if (sheet.PivotTables.Count == 0)
            {
                // Create sample data for the pivot table
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Product";
                cells["C1"].Value = "Sales";

                cells["A2"].Value = "Fruits";
                cells["B2"].Value = "Apple";
                cells["C2"].Value = 120;

                cells["A3"].Value = "Fruits";
                cells["B3"].Value = "Banana";
                cells["C3"].Value = 80;

                cells["A4"].Value = "Vegetables";
                cells["B4"].Value = "Carrot";
                cells["C4"].Value = 60;

                // Add a PivotTable based on the sample data
                int pivotIdx = sheet.PivotTables.Add("A1:C4", "E5", "DemoPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivot.RefreshData();
                pivot.CalculateData();
            }

            // Retrieve the first (or only) PivotTable on the sheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // -------------------------------------------------
            // 1. Create slicers
            // -------------------------------------------------
            SlicerCollection slicers = sheet.Slicers;

            // Add a slicer for the "Category" field, positioned at cell G5
            int slicerIdx1 = slicers.Add(pivotTable, "G5", "Category");
            Slicer slicer1 = slicers[slicerIdx1];

            // Add a second slicer for the "Product" field, positioned at cell G15
            int slicerIdx2 = slicers.Add(pivotTable, "G15", "Product");
            Slicer slicer2 = slicers[slicerIdx2];

            // -------------------------------------------------
            // 2. Modify slicer properties
            // -------------------------------------------------
            // Modify first slicer
            slicer1.Caption = "Category Filter";
            slicer1.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer1.NumberOfColumns = 2;
            slicer1.LockedPosition = true; // Prevent user from moving it
            slicer1.ShowCaption = true;
            slicer1.WidthPixel = 200;
            slicer1.HeightPixel = 120;

            // Modify second slicer
            slicer2.Caption = "Product Filter";
            slicer2.StyleType = SlicerStyleType.SlicerStyleLight1;
            slicer2.NumberOfColumns = 1;
            slicer2.LockedPosition = false; // Allow moving
            slicer2.ShowCaption = true;
            slicer2.WidthPixel = 180;
            slicer2.HeightPixel = 150;

            // Refresh slicers to ensure they reflect the latest PivotTable data
            slicer1.Refresh();
            slicer2.Refresh();

            // -------------------------------------------------
            // 3. Remove slicers
            // -------------------------------------------------
            // Example A: Remove the second slicer using RemoveAt (by index)
            slicers.RemoveAt(slicerIdx2); // After removal, slicerIdx1 remains valid

            // Example B: Remove the first slicer using Remove (by object reference)
            slicers.Remove(slicer1);

            // -------------------------------------------------
            // Save the modified workbook
            // -------------------------------------------------
            workbook.Save("Output.xlsx");
        }
    }
}