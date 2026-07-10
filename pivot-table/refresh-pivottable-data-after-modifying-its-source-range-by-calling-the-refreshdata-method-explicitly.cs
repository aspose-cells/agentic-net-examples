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

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(850);
            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["B4"].PutValue(300);
            sheet.Cells["A5"].PutValue("Banana");
            sheet.Cells["B5"].PutValue(400);

            // Add a pivot table based on the data range A1:B5, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Initial calculation to populate the pivot table
            pivotTable.CalculateData();

            // Modify the source data (e.g., change a sales value)
            sheet.Cells["B2"].PutValue(1500); // Updated sales for Apple

            // Refresh the pivot table data from the modified source range
            pivotTable.RefreshData();

            // Recalculate the pivot table after refresh
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTable_RefreshDemo.xlsx");
        }
    }
}