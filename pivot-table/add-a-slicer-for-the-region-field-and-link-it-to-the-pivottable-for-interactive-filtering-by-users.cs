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

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Apple";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 800;

            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 600;

            sheet.Cells["A5"].Value = "Banana";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 400;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivot.AddFieldToArea(PivotFieldType.Row, "Region");   // Region as row field
            pivot.AddFieldToArea(PivotFieldType.Column, "Product"); // Product as column field
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");   // Sales as data field

            // Refresh and calculate the pivot table so that slicer can work correctly
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer for the "Region" field; place it at cell G1
            int slicerIndex = sheet.Slicers.Add(pivot, "G1", "Region");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Link the slicer to the pivot table (required for interactive filtering)
            slicer.AddPivotConnection(pivot);

            // Optional: set a caption for better UI
            slicer.Caption = "Region Filter";

            // Save the workbook
            workbook.Save("PivotTableWithRegionSlicer.xlsx");
        }
    }
}