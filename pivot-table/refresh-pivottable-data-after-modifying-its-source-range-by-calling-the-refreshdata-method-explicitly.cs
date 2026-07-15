using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will serve as the pivot table source
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(850);
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(950);

            // Add a pivot table based on the source range A1:B4, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Initial calculation so the pivot shows correct values
            pivotTable.CalculateData();

            // ----- Modify the source data -----
            sheet.Cells["B2"].PutValue(1300); // Apple sales changed
            sheet.Cells["B3"].PutValue(900);  // Banana sales changed

            // Refresh the pivot table's cache from the updated source range
            // This explicitly calls RefreshData as required
            pivotTable.RefreshData();

            // Recalculate the pivot table to reflect refreshed data
            pivotTable.CalculateData();

            // Save the workbook (lifecycle: save)
            workbook.Save("PivotRefreshAfterSourceChange.xlsx");
        }
    }
}