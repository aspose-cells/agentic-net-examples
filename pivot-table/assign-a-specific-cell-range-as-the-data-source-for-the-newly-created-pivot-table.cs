using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class AssignPivotDataSource
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (A1:C5)
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Region");
        sheet.Cells["C1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue("North");
        sheet.Cells["C2"].PutValue(100);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue("South");
        sheet.Cells["C3"].PutValue(200);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue("East");
        sheet.Cells["C4"].PutValue(300);
        sheet.Cells["A5"].PutValue("D");
        sheet.Cells["B5"].PutValue("West");
        sheet.Cells["C5"].PutValue(400);

        // Add a pivot table with an initial source (A1:C5) at cell E3
        int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "MyPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");
        pivot.AddFieldToArea(PivotFieldType.Column, "Region");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Assign a new data source range (e.g., B2:C5) to the pivot table
        string[] newDataSource = new string[] { "B2:C5" };
        pivot.ChangeDataSource(newDataSource);

        // Refresh and recalculate to apply the new source
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook
        workbook.Save("PivotTableWithNewSource.xlsx");
    }
}