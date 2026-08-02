using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace AsposeCellsSlicerFontDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["A4"].Value = "Vegetable";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["B2"].Value = 10;
            sheet.Cells["B3"].Value = 20;
            sheet.Cells["B4"].Value = 15;

            // Add a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);      // Category field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);     // Amount field
            pivot.CalculateData();

            // Add a slicer linked to the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "F1", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Modify slicer label font: family, size, and color
            // The slicer label font is accessed via the Shape.Font property
            slicer.Shape.Font.Name = "Calibri";          // Font family
            slicer.Shape.Font.Size = 12;                 // Font size (points)
            slicer.Shape.Font.Color = Color.DarkBlue;    // Font color

            // Optional: make the caption visible and set its text
            slicer.ShowCaption = true;
            slicer.Caption = "Category Filter";

            // Save the workbook
            workbook.Save("SlicerFontDemo.xlsx");
        }
    }
}