using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Bike";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 5000;

            sheet.Cells["A3"].Value = "Bike";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 3000;

            sheet.Cells["A4"].Value = "Car";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 8000;

            sheet.Cells["A5"].Value = "Car";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 12000;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Calculate the pivot data
            pivotTable.CalculateData();

            // Assign the built‑in style PivotTableStyleMedium9
            pivotTable.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // Save the workbook (lifecycle save)
            workbook.Save("PivotTableWithMedium9Style.xlsx");
        }
    }
}