using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class RefreshPivotExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("Apple");
        sheet.Cells["B4"].PutValue(150);

        // Add a pivot table that uses the data range A1:B4 and place it starting at D3
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot table: Product as row field, Sales as data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column index 0 -> Product
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column index 1 -> Sales

        // Refresh the pivot cache from the source data
        pivotTable.RefreshData();

        // Calculate the pivot data into the worksheet cells
        pivotTable.CalculateData();

        // Save the workbook with the refreshed and calculated pivot table
        workbook.Save("RefreshPivotResult.xlsx");
    }
}