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

        // Populate sample data (Product and Sales)
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["A2"].PutValue("Confidential");
        sheet.Cells["A3"].PutValue("Public");
        sheet.Cells["A4"].PutValue("Confidential");
        sheet.Cells["A5"].PutValue("Public");

        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["B3"].PutValue(2000);
        sheet.Cells["B4"].PutValue(1500);
        sheet.Cells["B5"].PutValue(3000);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the Product field to the row area and Sales to the data area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Retrieve the row field (Product) to manipulate its items
        PivotField productField = pivotTable.RowFields[0];

        // Hide all items with the value "Confidential"
        // Using the HideItem(string, bool) overload as per the documentation
        productField.HideItem("Confidential", true);

        // Refresh the pivot table to apply changes and recalculate data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the hidden pivot items
        workbook.Save("HiddenPivotItemsDemo.xlsx");
    }
}