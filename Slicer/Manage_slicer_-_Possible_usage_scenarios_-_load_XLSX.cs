using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerManagementDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (lifecycle rule: use provided load method)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (or any specific worksheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is a pivot table; if not, create a simple one for demonstration
            PivotTable pivot;
            if (sheet.PivotTables.Count == 0)
            {
                // Create sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Vegetable");
                sheet.Cells["B3"].PutValue(80);

                // Add a pivot table covering the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B3", "D1", "DemoPivot");
                pivot = sheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            }
            else
            {
                // Use the first existing pivot table
                pivot = sheet.PivotTables[0];
            }

            // Add a slicer linked to the pivot table
            // Using the overload: Add(PivotTable pivot, string destCellName, string baseFieldName)
            int slicerIndex = sheet.Slicers.Add(pivot, "E2", "Category");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Configure slicer properties
            slicer.Caption = "Category Filter";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.NumberOfColumns = 1;
            slicer.LockedPosition = true; // Prevent moving/resizing via UI
            slicer.ShowCaption = true;
            slicer.ShowAllItems = true;

            // Refresh the slicer to ensure it reflects current pivot data
            slicer.Refresh();

            // Save the modified workbook (lifecycle rule: use provided save method)
            workbook.Save("output_with_slicer.xlsx", SaveFormat.Xlsx);
        }
    }
}