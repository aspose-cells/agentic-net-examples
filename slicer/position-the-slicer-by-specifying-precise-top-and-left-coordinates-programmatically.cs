using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace AsposeCellsSlicerPositionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with sample data for a pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["A3"].Value = "Vegetable";
            sheet.Cells["B1"].Value = "Sales";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["B3"].Value = 80;

            // Add a pivot table based on the sample data
            int pivotIdx = sheet.PivotTables.Add("A1:B3", "D5", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer for the "Category" field of the pivot table
            // The slicer will be placed initially at cell E5 (upper‑left corner)
            int slicerIdx = sheet.Slicers.Add(pivot, "E5", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Position the slicer precisely using its underlying Shape object
            // Values are in pixels relative to the worksheet's top‑left corner
            slicer.Shape.Left = 150;   // Horizontal offset from the left edge
            slicer.Shape.Top = 80;     // Vertical offset from the top edge

            // Optionally set size (also using the Shape object)
            slicer.Shape.Width = 200;
            slicer.Shape.Height = 150;

            // Save the workbook
            workbook.Save("SlicerPositionedDemo.xlsx");
        }
    }
}