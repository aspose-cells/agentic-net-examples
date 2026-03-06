using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Apple");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(800);

            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue("North");
            sheet.Cells["C4"].PutValue(1500);

            sheet.Cells["A5"].PutValue("Banana");
            sheet.Cells["B5"].PutValue("South");
            sheet.Cells["C5"].PutValue(700);

            // Add a pivot table based on the source data
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row, Region as column, Sales as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Add a calculated field (e.g., SalesTax = Sales * 0.1)
            // The second parameter is the formula using field names
            pivotTable.AddCalculatedField("SalesTax", "Sales*0.1");

            // Refresh the pivot cache and recalculate the pivot table
            pivotTable.RefreshData();          // Refreshes data from the source
            pivotTable.CalculateData();        // Calculates the pivot results, including the calculated field

            // Optionally, refresh all pivot tables in the workbook (not required here but shown for completeness)
            // workbook.Worksheets.RefreshPivotTables();

            // Save the workbook in XLSX format
            workbook.Save("PivotTable_With_CalculatedField.xlsx", SaveFormat.Xlsx);
        }
    }
}