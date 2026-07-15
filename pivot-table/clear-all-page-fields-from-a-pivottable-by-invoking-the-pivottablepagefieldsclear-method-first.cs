using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class ClearPivotTablePageFields
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Region");
        sheet.Cells["B1"].PutValue("Product");
        sheet.Cells["C1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("North");
        sheet.Cells["B2"].PutValue("Apple");
        sheet.Cells["C2"].PutValue(100);
        sheet.Cells["A3"].PutValue("South");
        sheet.Cells["B3"].PutValue("Banana");
        sheet.Cells["C3"].PutValue(150);

        // Add a pivot table; include a page field ("Region")
        int pivotIndex = sheet.PivotTables.Add("A1:C3", "E5", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivotTable.AddFieldToArea(PivotFieldType.Page, "Region");

        // Build the pivot table
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Clear all page fields from the pivot table
        pivotTable.PageFields.Clear();

        // Rebuild the pivot after clearing page fields
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("ClearPageFieldsDemo.xlsx");
    }
}