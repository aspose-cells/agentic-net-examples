using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    class Program
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
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["B4"].PutValue(150);

            // Add a pivot table based on the source data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Initial calculation so the pivot shows correct values
            pivot.RefreshData();
            pivot.CalculateData();

            // ----- Change the source data -----
            sheet.Cells["B2"].PutValue(200); // Update Apple sales
            sheet.Cells["B3"].PutValue(90);  // Update Banana sales

            // Refresh only the specific pivot table to reflect the changes
            pivot.RefreshData();   // Refreshes data from the source
            pivot.CalculateData(); // Recalculates the pivot report

            // Save the workbook
            workbook.Save("PivotRefreshResult.xlsx");
        }
    }
}