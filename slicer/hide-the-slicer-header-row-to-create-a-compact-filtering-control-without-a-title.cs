using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerHeaderHideDemo
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
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Fruit";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Vegetable";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Grain";
            cells["B4"].Value = 150;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D5", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.CalculateData();

            // Add a slicer linked to the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "F5", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Hide the slicer header (caption) to make it compact
            slicer.ShowCaption = false;

            // Optional: adjust appearance
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;
            slicer.NumberOfColumns = 1;

            // Save the workbook
            workbook.Save("SlicerHeaderHidden.xlsx");
        }
    }
}