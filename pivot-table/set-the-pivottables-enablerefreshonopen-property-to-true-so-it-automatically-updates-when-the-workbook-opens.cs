using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshOnOpenDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1000);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(2000);
            worksheet.Cells["A4"].PutValue("Orange");
            worksheet.Cells["B4"].PutValue(1500);

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure the pivot table (e.g., add row and data fields)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column

            // Enable automatic refresh when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = true;

            // Save the workbook to a file
            workbook.Save("PivotTableRefreshOnOpen.xlsx");
        }
    }
}