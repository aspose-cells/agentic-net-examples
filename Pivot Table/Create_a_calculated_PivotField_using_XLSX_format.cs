using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCalculatedPivotFieldDemo
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
            // Header row
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Region";
            cells["C1"].Value = "Sales";

            // Data rows
            cells["A2"].Value = "Apple";   cells["B2"].Value = "North"; cells["C2"].Value = 1200;
            cells["A3"].Value = "Apple";   cells["B3"].Value = "South"; cells["C3"].Value = 800;
            cells["A4"].Value = "Banana";  cells["B4"].Value = "North"; cells["C4"].Value = 1500;
            cells["A5"].Value = "Banana";  cells["B5"].Value = "South"; cells["C5"].Value = 700;
            cells["A6"].Value = "Cherry";  cells["B6"].Value = "North"; cells["C6"].Value = 900;
            cells["A7"].Value = "Cherry";  cells["B7"].Value = "South"; cells["C7"].Value = 1100;

            // Add a pivot table covering the data range A1:C7, place it at E3, and name it "SalesPivot"
            int pivotIndex = sheet.PivotTables.Add("A1:C7", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");   // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region"); // Column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");    // Data field

            // Add a calculated field named "DoubleSales" with formula "=Sales*2"
            // The third parameter 'true' drags the field to the data area automatically
            pivotTable.AddCalculatedField("DoubleSales", "=Sales*2", true);

            // Refresh the pivot cache and calculate the pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("CalculatedPivotFieldDemo.xlsx");
        }
    }
}