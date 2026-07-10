using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotFieldHideDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Region";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Apple";
            cells["B2"].Value = "North";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "Apple";
            cells["B3"].Value = "South";
            cells["C3"].Value = 800;

            cells["A4"].Value = "Banana";
            cells["B4"].Value = "North";
            cells["C4"].Value = 1500;

            cells["A5"].Value = "Banana";
            cells["B5"].Value = "South";
            cells["C5"].Value = 700;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");   // Row field we will hide later
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region"); // Column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");    // Data field

            // Hide the "Product" field from the report area.
            // This is achieved by removing the field from its current area (Row).
            // The RemoveField method matches the available rule.
            pivotTable.RemoveField(PivotFieldType.Row, "Product");

            // Refresh and calculate the pivot table after modification
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotFieldHiddenDemo.xlsx");
        }
    }
}