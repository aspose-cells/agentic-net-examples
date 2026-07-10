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

        // Populate sample data (Product | Sales)
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");

        string[] products = { "Public1", "ConfidentialA", "Public2", "ConfidentialB", "Public3" };
        int[] sales = { 1000, 2000, 1500, 3000, 1200 };

        for (int i = 0; i < products.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(products[i]); // Column A
            sheet.Cells[i + 2, 1].PutValue(sales[i]);   // Column B
        }

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add row field (Product) and data field (Sales) to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Get the row pivot field that contains product names
        PivotField productField = pivotTable.RowFields[0];

        // Define confidential items that must be hidden from the report view
        string[] confidentialItems = { "ConfidentialA", "ConfidentialB" };

        // Hide each confidential item using the HideItem(string, bool) method
        foreach (string item in confidentialItems)
        {
            productField.HideItem(item, true);
        }

        // Refresh and calculate the pivot table to apply changes
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with hidden pivot items
        workbook.Save("HiddenPivotItemsDemo.xlsx");
    }
}