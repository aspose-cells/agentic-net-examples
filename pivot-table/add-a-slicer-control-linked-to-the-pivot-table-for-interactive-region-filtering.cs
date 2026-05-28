using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with a Region field (used for the slicer)
            cells["A1"].Value = "Region";
            cells["B1"].Value = "Product";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "North";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 120;

            cells["A3"].Value = "South";
            cells["B3"].Value = "Banana";
            cells["C3"].Value = 150;

            cells["A4"].Value = "North";
            cells["B4"].Value = "Orange";
            cells["C4"].Value = 200;

            cells["A5"].Value = "East";
            cells["B5"].Value = "Apple";
            cells["C5"].Value = 130;

            cells["A6"].Value = "West";
            cells["B6"].Value = "Banana";
            cells["C6"].Value = 170;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C6", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot: Region as row field, Product as column field, Sales as data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Region");
            pivot.AddFieldToArea(PivotFieldType.Column, "Product");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table for the "Region" field
            // The slicer will be placed with its upper‑left corner at cell G3
            int slicerIndex = sheet.Slicers.Add(pivot, "G3", "Region");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Optional: customize slicer appearance
            slicer.Caption = "Region Filter";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.NumberOfColumns = 1;
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 120;

            // Save the workbook
            workbook.Save("SlicerWithPivot.xlsx");
        }
    }
}