using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Bike";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 5000;

            sheet.Cells["A3"].Value = "Car";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 12000;

            sheet.Cells["A4"].Value = "Truck";
            sheet.Cells["B4"].Value = "East";
            sheet.Cells["C4"].Value = 8000;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Calculate the pivot data
            pivotTable.CalculateData();

            // Assign the built‑in style PivotTableStyleMedium9
            pivotTable.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // Save the workbook
            workbook.Save("PivotTableWithMedium9Style.xlsx");
        }
    }
}