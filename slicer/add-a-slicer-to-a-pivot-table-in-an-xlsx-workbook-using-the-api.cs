using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
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
            cells["C4"].Value = 150;

            cells["A5"].Value = "Vegetables";
            cells["B5"].Value = "Tomato";
            cells["C5"].Value = 200;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E2", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "Product");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table for the "Category" field
            // The slicer will be placed with its upper‑left corner at cell G2
            int slicerIndex = sheet.Slicers.Add(pivot, "G2", "Category");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Optional: set some slicer properties
            slicer.Caption = "Category Filter";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.NumberOfColumns = 1;

            // Save the workbook to an XLSX file
            workbook.Save("PivotWithSlicer.xlsx", SaveFormat.Xlsx);
        }
    }
}