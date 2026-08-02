using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class RemoveCalculatedFieldDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["C1"].PutValue("Quantity");

        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["C2"].PutValue(5);

        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["C3"].PutValue(8);

        sheet.Cells["A4"].PutValue("Orange");
        sheet.Cells["B4"].PutValue(200);
        sheet.Cells["C4"].PutValue(10);

        // Add a pivot table covering the data range and place it at E3
        int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table: Product as row, Sales and Quantity as data
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

        // Add a calculated field named "Total" (Sales * Quantity) and drag it to the data area
        pivotTable.AddCalculatedField("Total", "=Sales*Quantity", true);

        // Refresh and calculate to populate the pivot table with the calculated field
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Remove the calculated field "Total" from the Data area without affecting other fields
        pivotTable.RemoveField(PivotFieldType.Data, "Total");

        // Recalculate after removal to update the pivot view
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("RemovedCalculatedField.xlsx");
    }
}