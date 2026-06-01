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

            sheet.Cells["A2"].Value = "Laptop";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Laptop";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 800;

            sheet.Cells["A4"].Value = "Desktop";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 1500;

            sheet.Cells["A5"].Value = "Desktop";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 1100;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Apply a predefined built‑in pivot table style for consistent formatting
            // Using the PivotTableStyleName property (available rule)
            pivotTable.PivotTableStyleName = "PivotStyleLight16";

            // Save the workbook with the styled pivot table
            workbook.Save("PivotTableWithPredefinedStyle.xlsx", SaveFormat.Xlsx);
        }
    }
}