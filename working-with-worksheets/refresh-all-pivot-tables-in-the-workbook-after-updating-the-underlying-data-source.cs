using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(3000);

            // Add a pivot table based on the source data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

            // Modify the underlying data source
            sheet.Cells["B2"].PutValue(1500); // Apple sales changed
            sheet.Cells["B3"].PutValue(2500); // Orange sales changed

            // Refresh all pivot tables in the workbook to reflect the changes
            workbook.Worksheets.RefreshPivotTables();

            // Save the updated workbook
            workbook.Save("RefreshedPivotTables.xlsx");
        }
    }
}