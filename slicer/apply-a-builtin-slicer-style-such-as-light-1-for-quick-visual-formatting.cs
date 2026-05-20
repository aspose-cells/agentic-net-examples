using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace AsposeCellsSlicerStyleDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(300);

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D2", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

            // Add a slicer linked to the pivot table's first base field (Category)
            int slicerIdx = sheet.Slicers.Add(pivot, "F2", 0);
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Apply a built‑in slicer style (Light 1)
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Optional: set additional slicer properties for better visibility
            slicer.Caption = "Category Filter";
            slicer.NumberOfColumns = 1;
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 100;

            // Save the workbook with the styled slicer
            workbook.Save("SlicerStyleLight1Demo.xlsx");
        }
    }
}