using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshOnOpenDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(850);
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(950);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table (row field and data field)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sales column

            // Enable auto‑refresh when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = true;

            // Save the workbook
            workbook.Save("PivotTable_AutoRefreshOnOpen.xlsx");
        }
    }
}