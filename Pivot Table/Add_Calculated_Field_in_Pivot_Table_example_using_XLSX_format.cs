using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCalculatedFieldExample
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
            sheet.Cells["C1"].PutValue("Quantity");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["C2"].PutValue(5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["C3"].PutValue(8);

            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["C4"].PutValue(10);

            // Add a pivot table based on the data range A1:C4, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Sales and Quantity as data fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Add a calculated field that computes total revenue = Sales * Quantity
            // The third parameter 'true' drags the field to the data area automatically
            pivotTable.AddCalculatedField("TotalRevenue", "=Sales*Quantity", true);

            // Refresh and calculate the pivot table to ensure the calculated field is evaluated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("PivotTableWithCalculatedField.xlsx");
        }
    }
}