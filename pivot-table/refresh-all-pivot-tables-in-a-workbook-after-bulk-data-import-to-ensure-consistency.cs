using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet for data import
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Simulate bulk data import by filling a range with sample values
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            for (int i = 2; i <= 101; i++)
            {
                // Alternate categories "A" and "B"
                dataSheet.Cells[$"A{i}"].PutValue(i % 2 == 0 ? "A" : "B");
                // Populate amount with a simple calculation
                dataSheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a new worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create a pivot table based on the imported data range
            // Note: the range includes the sheet name ("Data!") to reference the source worksheet
            int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:B101", "D3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows = Category, data = Sum of Amount
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Amount

            // Refresh all pivot tables in the workbook to reflect the newly imported data
            workbook.Worksheets.RefreshPivotTables();

            // Save the workbook with the refreshed pivot table
            workbook.Save("PivotRefreshAfterImport.xlsx");
        }
    }
}