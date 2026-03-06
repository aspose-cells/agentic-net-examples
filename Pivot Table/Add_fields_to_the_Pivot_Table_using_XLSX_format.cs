using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Region");
        sheet.Cells["C1"].PutValue("Sales");

        sheet.Cells["A2"].PutValue("P1");
        sheet.Cells["B2"].PutValue("North");
        sheet.Cells["C2"].PutValue(1000);

        sheet.Cells["A3"].PutValue("P2");
        sheet.Cells["B3"].PutValue("South");
        sheet.Cells["C3"].PutValue(1500);

        sheet.Cells["A4"].PutValue("P3");
        sheet.Cells["B4"].PutValue("East");
        sheet.Cells["C4"].PutValue(2000);

        sheet.Cells["A5"].PutValue("P1");
        sheet.Cells["B5"].PutValue("West");
        sheet.Cells["C5"].PutValue(1200);

        // Add a pivot table (source range, destination cell, table name)
        int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table using the string overload of AddFieldToArea
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");   // Row area
        pivot.AddFieldToArea(PivotFieldType.Column, "Region"); // Column area
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");    // Data area

        // Refresh and calculate the pivot table to populate data
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("PivotTableWithFields.xlsx");
    }
}