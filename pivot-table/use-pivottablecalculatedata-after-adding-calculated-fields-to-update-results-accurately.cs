using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCalculatedFieldDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            // Columns: Product, Quantity, Price
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(2);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(5);
            sheet.Cells["C3"].PutValue(3);

            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(8);
            sheet.Cells["C4"].PutValue(1.5);

            // Add a pivot table based on the data range A1:C4, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row, Quantity and Price as data fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Price");

            // Add a calculated field that computes total sales = Quantity * Price
            // The formula uses the field names exactly as they appear in the source data
            pivotTable.AddCalculatedField("TotalSales", "=Quantity*Price", true);

            // Refresh the pivot cache to capture any changes in the source data
            pivotTable.RefreshData();

            // Calculate the pivot data so that the calculated field values are populated
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTable_With_CalculatedField.xlsx");
        }
    }
}