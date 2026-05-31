using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: Product categories and their sales
        sheet.Cells["A1"].Value = "Product";
        sheet.Cells["B1"].Value = "Sales";

        string[] products = { "Apple", "Banana", "Orange", "Grapes", "Mango" };
        int[] sales = { 1000, 2000, 3000, 4000, 5000 };

        for (int i = 0; i < products.Length; i++)
        {
            sheet.Cells[i + 2, 0].Value = products[i];
            sheet.Cells[i + 2, 1].Value = sales[i];
        }

        // Create a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the Product field to the row area and Sales to the data area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Retrieve the row field (Product) to which we will add a calculated item
        PivotField productField = pivotTable.RowFields[0];

        // Add a calculated item that groups Apple and Banana into a custom group named "FruitGroup"
        // The formula uses the names of the base items as they appear in the source data
        productField.AddCalculatedItem("FruitGroup", "=Apple + Banana");

        // Refresh and calculate the pivot table to reflect the new calculated item
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the pivot table and calculated group
        workbook.Save("PivotCalculatedGroup.xlsx");
    }
}