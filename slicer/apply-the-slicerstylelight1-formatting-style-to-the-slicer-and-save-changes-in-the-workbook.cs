using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerStyleDemo
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
            cells["A2"].Value = "Apple";
            cells["A3"].Value = "Orange";
            cells["A4"].Value = "Banana";
            cells["B1"].Value = "Quantity";
            cells["B2"].Value = 10;
            cells["B3"].Value = 15;
            cells["B4"].Value = 20;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0); // Fruit field

            // Add a slicer linked to the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "A1", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Apply the built‑in light style 1 to the slicer
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Save the workbook with the styled slicer
            workbook.Save("SlicerStyleLight1Demo.xlsx");
        }
    }
}