using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample sales data
            // Columns: Region, Category, Sales
            sheet.Cells["A1"].PutValue("Region");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["C1"].PutValue("Sales");

            // Sample rows
            sheet.Cells["A2"].PutValue("North");
            sheet.Cells["B2"].PutValue("Widgets");
            sheet.Cells["C2"].PutValue(5000);

            sheet.Cells["A3"].PutValue("North");
            sheet.Cells["B3"].PutValue("Gadgets");
            sheet.Cells["C3"].PutValue(3000);

            sheet.Cells["A4"].PutValue("South");
            sheet.Cells["B4"].PutValue("Widgets");
            sheet.Cells["C4"].PutValue(6000);

            sheet.Cells["A5"].PutValue("South");
            sheet.Cells["B5"].PutValue("Gadgets");
            sheet.Cells["C5"].PutValue(4000);

            sheet.Cells["A6"].PutValue("West");
            sheet.Cells["B6"].PutValue("Widgets");
            sheet.Cells["C6"].PutValue(4500);

            // Define the source range for the pivot table (including headers)
            string sourceRange = "A1:C6";

            // Add a pivot table at cell E3
            int pivotIndex = sheet.PivotTables.Add(sourceRange, "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add row fields to group by Region and Category
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the Sales field as a data field (default aggregation is Sum)
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table to populate data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("GroupedSalesPivot.xlsx");
        }
    }
}