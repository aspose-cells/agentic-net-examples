using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class PivotGrandTotalRowsOnly
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Region");
        sheet.Cells["C1"].PutValue("Sales");

        sheet.Cells["A2"].PutValue("Product A");
        sheet.Cells["B2"].PutValue("North");
        sheet.Cells["C2"].PutValue(1000);

        sheet.Cells["A3"].PutValue("Product B");
        sheet.Cells["B3"].PutValue("South");
        sheet.Cells["C3"].PutValue(1500);

        sheet.Cells["A4"].PutValue("Product A");
        sheet.Cells["B4"].PutValue("South");
        sheet.Cells["C4"].PutValue(2000);

        sheet.Cells["A5"].PutValue("Product B");
        sheet.Cells["B5"].PutValue("North");
        sheet.Cells["C5"].PutValue(1200);

        // Add a pivot table based on the data range A1:C5, place it at E3, and name it "PivotTable1"
        PivotTableCollection pivotTables = sheet.PivotTables;
        int pivotIndex = pivotTables.Add("A1:C5", "E3", "PivotTable1");
        PivotTable pivotTable = pivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product as row field
        pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Region as column field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales as data field

        // Show grand totals for rows only
        pivotTable.ShowRowGrandTotals = true;    // Enable row grand totals
        pivotTable.ShowColumnGrandTotals = false; // Disable column grand totals

        // Save the workbook to a file
        workbook.Save("PivotTableRowGrandTotalsOnly.xlsx");
    }
}