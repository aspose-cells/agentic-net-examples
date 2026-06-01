using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a "Region" field
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Banana";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 850;

            sheet.Cells["A4"].Value = "Apple";
            sheet.Cells["B4"].Value = "South";
            sheet.Cells["C4"].Value = 950;

            sheet.Cells["A5"].Value = "Banana";
            sheet.Cells["B5"].Value = "North";
            sheet.Cells["C5"].Value = 1100;

            // Add a PivotTable based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the PivotTable: Product as Row, Region as Column, Sales as Data
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Column, "Region");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the PivotTable so that slicer can work correctly
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer for the "Region" field; place it at cell G1
            int slicerIndex = sheet.Slicers.Add(pivot, "G1", "Region");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Explicitly link the slicer to the PivotTable (optional, but ensures connection)
            slicer.AddPivotConnection(pivot);

            // Optionally set some slicer properties
            slicer.Caption = "Region Filter";
            slicer.NumberOfColumns = 1;
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 100;

            // Save the workbook
            workbook.Save("PivotTableWithRegionSlicer.xlsx");
        }
    }
}