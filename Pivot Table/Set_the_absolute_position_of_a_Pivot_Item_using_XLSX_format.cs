using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class SetPivotItemAbsolutePosition
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["A4"].PutValue("Orange");
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["B3"].PutValue(2000);
        sheet.Cells["B4"].PutValue(3000);

        // Add a pivot table covering the data range and place it at E3
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the row field (Product) and the data field (Sales)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Refresh and calculate to generate pivot items
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Access the row field containing the product items
        PivotField productField = pivotTable.RowFields[0];

        // Set absolute positions for specific pivot items
        // Position is the index among all pivot items (not limited to the same parent node)
        PivotItem bananaItem = productField.PivotItems["Banana"];
        bananaItem.Position = 0; // Move "Banana" to the first position

        PivotItem orangeItem = productField.PivotItems["Orange"];
        orangeItem.Position = 1; // Move "Orange" to the second position

        // Recalculate after changing positions
        pivotTable.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("PivotItemAbsolutePosition.xlsx");
    }
}