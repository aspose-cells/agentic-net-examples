using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Fill sample data that will be used for the pivot table
            // -------------------------------------------------
            Cells cells = sheet.Cells;
            cells["A1"].PutValue("Product");   // Row field
            cells["B1"].PutValue("Region");    // Column field
            cells["C1"].PutValue("Category");  // Page field
            cells["D1"].PutValue("Sales");     // Data field

            // Add a few rows of data
            cells["A2"].PutValue("Bike");   cells["B2"].PutValue("North"); cells["C2"].PutValue("Sports");   cells["D2"].PutValue(1200);
            cells["A3"].PutValue("Bike");   cells["B3"].PutValue("South"); cells["C3"].PutValue("Sports");   cells["D3"].PutValue(800);
            cells["A4"].PutValue("Car");    cells["B4"].PutValue("North"); cells["C4"].PutValue("Luxury");   cells["D4"].PutValue(2500);
            cells["A5"].PutValue("Car");    cells["B5"].PutValue("South"); cells["C5"].PutValue("Luxury");   cells["D5"].PutValue(3000);
            cells["A6"].PutValue("Truck");  cells["B6"].PutValue("North"); cells["C6"].PutValue("Utility");  cells["D6"].PutValue(1500);
            cells["A7"].PutValue("Truck");  cells["B7"].PutValue("South"); cells["C7"].PutValue("Utility");  cells["D7"].PutValue(1800);

            // -------------------------------------------------
            // 2. Create a pivot table
            // -------------------------------------------------
            // Define the source data range (including headers)
            string sourceData = "A1:D7";

            // Destination cell where the pivot table will start
            string destCell = "F3";

            // Name of the pivot table
            string pivotName = "SalesPivot";

            // Add the pivot table to the worksheet
            int pivotIndex = sheet.PivotTables.Add(sourceData, destCell, pivotName);
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // -------------------------------------------------
            // 3. Add fields to the appropriate areas
            // -------------------------------------------------
            // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Column field
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");

            // Data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Page (Report Filter) field
            // The PivotFieldType.Page enum value is supported for adding page fields
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Category");

            // -------------------------------------------------
            // 4. Refresh and calculate the pivot table data
            // -------------------------------------------------
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // 5. Save the workbook in XLSX format
            // -------------------------------------------------
            workbook.Save("PivotTableWithAllFields.xlsx", SaveFormat.Xlsx);
        }
    }
}