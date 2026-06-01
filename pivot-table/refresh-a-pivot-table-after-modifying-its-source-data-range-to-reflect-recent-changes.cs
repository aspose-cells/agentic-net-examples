using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class RefreshPivotDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        dataSheet.Cells["A1"].PutValue("Product");
        dataSheet.Cells["B1"].PutValue("Sales");
        dataSheet.Cells["A2"].PutValue("Apple");
        dataSheet.Cells["B2"].PutValue(100);
        dataSheet.Cells["A3"].PutValue("Banana");
        dataSheet.Cells["B3"].PutValue(200);
        dataSheet.Cells["A4"].PutValue("Apple");
        dataSheet.Cells["B4"].PutValue(150);

        // Add a pivot table based on the source range
        int pivotIndex = dataSheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
        PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];

        // Configure the pivot table: Product as row field, Sales as data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Product
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Sales

        // Initial calculation to populate the pivot table
        pivotTable.CalculateData();

        // ----- Modify the source data after the pivot table has been created -----
        dataSheet.Cells["B2"].PutValue(120); // Update Apple sales
        dataSheet.Cells["B3"].PutValue(250); // Update Banana sales
        dataSheet.Cells["A5"].PutValue("Orange"); // Add a new product
        dataSheet.Cells["B5"].PutValue(180); // Sales for Orange

        // Refresh the pivot cache to reflect the changed source data
        pivotTable.RefreshData();

        // Recalculate the pivot table so the new data appears in the worksheet
        pivotTable.CalculateData();

        // Save the workbook with the refreshed pivot table
        workbook.Save("RefreshedPivot.xlsx");
    }
}