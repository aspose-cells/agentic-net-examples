using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class RefreshCalculatePivotWithCalculatedItems
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Product");
        sheet.Cells["C1"].PutValue("Sales");

        sheet.Cells["A2"].PutValue("Fruit");
        sheet.Cells["B2"].PutValue("Apple");
        sheet.Cells["C2"].PutValue(120);

        sheet.Cells["A3"].PutValue("Fruit");
        sheet.Cells["B3"].PutValue("Banana");
        sheet.Cells["C3"].PutValue(80);

        sheet.Cells["A4"].PutValue("Vegetable");
        sheet.Cells["B4"].PutValue("Carrot");
        sheet.Cells["C4"].PutValue(50);

        sheet.Cells["A5"].PutValue("Vegetable");
        sheet.Cells["B5"].PutValue("Broccoli");
        sheet.Cells["C5"].PutValue(70);

        // Add a pivot table based on the source data
        int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Add a calculated field (e.g., AdjustedSales = Sales * 1.1) and place it in the data area
        pivotTable.AddCalculatedField("AdjustedSales", "Sales*1.1");

        // Refresh the pivot cache from the source data and calculate the pivot table values
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Refresh all pivot tables in the workbook (optional, ensures consistency)
        workbook.Worksheets.RefreshPivotTables();

        // Save the workbook in XLSX format
        workbook.Save("RefreshCalculatePivotWithCalculatedItems.xlsx");
    }
}