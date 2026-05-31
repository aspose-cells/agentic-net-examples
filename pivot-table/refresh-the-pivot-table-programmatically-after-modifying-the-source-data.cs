using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotRefreshExample
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
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["B4"].PutValue(500);

            // Add a pivot table based on the source range A1:B4, place it at C3
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "C3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Initial calculation so the pivot shows data
            pivotTable.CalculateData();

            // ----- Modify the source data -----
            sheet.Cells["B2"].PutValue(1500); // Update Apple sales
            sheet.Cells["A5"].PutValue("Banana"); // Add a new row
            sheet.Cells["B5"].PutValue(600);

            // Refresh the pivot table to reflect the changes
            // RefreshData gathers data from the source into the pivot cache
            // CalculateData writes the refreshed data into the worksheet cells
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the refreshed pivot table
            workbook.Save("PivotTable_Refreshed.xlsx");
        }
    }
}