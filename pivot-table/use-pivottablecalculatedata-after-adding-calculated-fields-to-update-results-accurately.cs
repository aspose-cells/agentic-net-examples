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
            // Columns: Product, Quantity, UnitPrice
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("UnitPrice");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(2.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(15);
            sheet.Cells["C3"].PutValue(1.8);

            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(8);
            sheet.Cells["C4"].PutValue(3.0);

            // Add a pivot table based on the data range A1:C4, place it at D3
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Quantity and UnitPrice as data fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "UnitPrice");

            // Add a calculated field named "TotalSales" with formula =Quantity*UnitPrice
            // The third parameter 'true' drags the field to the data area automatically
            pivotTable.AddCalculatedField("TotalSales", "=Quantity*UnitPrice", true);

            // Refresh the pivot cache from the source data and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTable_With_CalculatedField.xlsx");
        }
    }
}