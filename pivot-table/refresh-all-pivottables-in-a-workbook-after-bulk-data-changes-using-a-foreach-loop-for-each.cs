using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class RefreshPivotTablesDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet and add sample data
        Worksheet ws = workbook.Worksheets[0];
        ws.Cells["A1"].PutValue("Product");
        ws.Cells["B1"].PutValue("Sales");
        ws.Cells["A2"].PutValue("Apple");
        ws.Cells["B2"].PutValue(1000);
        ws.Cells["A3"].PutValue("Orange");
        ws.Cells["B3"].PutValue(2000);

        // Add a pivot table based on the sample data
        int pivotIndex = ws.PivotTables.Add("A1:B3", "D5", "SalesPivot");
        PivotTable pivotTable = ws.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column as data field

        // Perform bulk data changes
        ws.Cells["B2"].PutValue(1500); // Update Apple sales
        ws.Cells["B3"].PutValue(2500); // Update Orange sales

        // Refresh all pivot tables in each worksheet using a foreach loop
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.RefreshPivotTables();
        }

        // Save the updated workbook
        workbook.Save("RefreshedPivotTables.xlsx");
    }
}