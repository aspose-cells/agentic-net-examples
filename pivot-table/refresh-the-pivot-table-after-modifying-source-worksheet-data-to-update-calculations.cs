using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (will hold source data)
            Worksheet dataSheet = workbook.Worksheets[0];
            Cells cells = dataSheet.Cells;

            // Populate source data for the pivot table
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("Fruit");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("Fruit");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("Vegetable");
            cells["B4"].PutValue(15);

            // Add a new worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create a pivot table based on the source range A1:B4
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B4", "C3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table (row field and data field)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Initial calculation so the pivot table shows data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ----- Modify source data -----
            cells["B2"].PutValue(30); // Change Fruit amount from 10 to 30
            cells["B4"].PutValue(25); // Change Vegetable amount from 15 to 25

            // Refresh all pivot tables in the worksheet containing the pivot table
            // This updates the pivot cache and recalculates the pivot report
            pivotSheet.RefreshPivotTables();

            // Alternatively, you could refresh all pivot tables in the entire workbook:
            // workbook.Worksheets.RefreshPivotTables();

            // Save the updated workbook
            workbook.Save("RefreshedPivotTableDemo.xlsx");
        }
    }
}