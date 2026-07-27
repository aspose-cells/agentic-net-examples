using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["B4"].PutValue(300);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the "Product" field to the row area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

        // Refresh and calculate the pivot table to populate items
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Get the collection of pivot items for the row field
        PivotItemCollection items = pivotTable.RowFields[0].PivotItems;

        // Move the first item two positions forward (down) within the same parent node
        // count = 2 (positive moves down), isSameParent = true
        if (items.Count > 1)
        {
            items[0].Move(2, true);
        }

        // Save the workbook with the modified pivot item order
        wb.Save("PivotItemMoveTwoPositions.xlsx");
    }
}