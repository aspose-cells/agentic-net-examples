using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace SlicerPositionExample
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
            sheet.Cells["B1"].Value = "Sales";
            sheet.Cells["A2"].Value = "Food";
            sheet.Cells["B2"].Value = 1200;
            sheet.Cells["A3"].Value = "Beverage";
            sheet.Cells["B3"].Value = 800;
            sheet.Cells["A4"].Value = "Electronics";
            sheet.Cells["B4"].Value = 1500;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D2", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Add a slicer linked to the pivot table for the "Category" field
            // Destination cell "F2" is the upper‑left corner of the slicer range
            int slicerIdx = sheet.Slicers.Add(pivot, "F2", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Position the slicer precisely using the Shape object (pixel units)
            // Set the left offset (horizontal) to 100 pixels from the worksheet's left edge
            slicer.Shape.Left = 100;
            // Set the top offset (vertical) to 50 pixels from the worksheet's top edge
            slicer.Shape.Top = 50;

            // Optionally, adjust size if needed
            slicer.Shape.Width = 200;   // width in pixels
            slicer.Shape.Height = 150;  // height in pixels

            // Save the workbook
            workbook.Save("SlicerPositioned.xlsx");
        }
    }
}