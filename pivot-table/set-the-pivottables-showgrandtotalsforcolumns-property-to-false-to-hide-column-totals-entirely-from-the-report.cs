using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
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
            sheet.Cells["C3"].Value = 1500;

            sheet.Cells["A4"].Value = "Phone";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 800;

            sheet.Cells["A5"].Value = "Phone";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 950;

            // Add a pivot table covering the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows = Product, columns = Region, data = Sales
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Region
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

            // Hide column grand totals (ShowGrandTotalsForColumns = false)
            pivotTable.ShowColumnGrandTotals = false;

            // Recalculate the pivot table after changing the setting
            pivotTable.CalculateData();

            // Save the workbook (lifecycle: save)
            workbook.Save("PivotTable_NoColumnGrandTotals.xlsx");
        }
    }
}