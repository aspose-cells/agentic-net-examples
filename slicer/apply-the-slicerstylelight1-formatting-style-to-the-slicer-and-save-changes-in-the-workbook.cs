using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Fruit";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B1"].Value = "Quantity";
            sheet.Cells["B2"].Value = 10;
            sheet.Cells["B3"].Value = 15;
            sheet.Cells["B4"].Value = 20;

            // Add a pivot table based on the sample data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Add "Fruit" as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Add "Quantity" as data field

            // Add a slicer linked to the pivot table's first field ("Fruit")
            int slicerIdx = sheet.Slicers.Add(pivot, "A1", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Apply the built‑in light style 1 to the slicer
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Save the workbook with the styled slicer
            workbook.Save("SlicerStyleLight1Demo.xlsx");
        }
    }
}