using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class HidePivotItemsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["A5"].PutValue("A");
        sheet.Cells["B5"].PutValue(40);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the row field (Category) and the data field (Value)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

        // Populate the pivot table
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Hide all pivot items except the one named "A"
        PivotField rowField = pivotTable.RowFields[0];
        foreach (PivotItem item in rowField.PivotItems)
        {
            // Setting IsHidden to true hides the item; false shows it
            item.IsHidden = (item.Name != "A");
        }

        // Recalculate after changing visibility
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("HiddenPivotItemsDemo.xlsx");
    }
}