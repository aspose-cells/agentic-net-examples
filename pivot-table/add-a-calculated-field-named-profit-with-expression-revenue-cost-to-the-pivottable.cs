using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCalculatedFieldDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample source data for the pivot table
            // Columns: Product, Revenue, Cost
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Revenue");
            sheet.Cells["C1"].PutValue("Cost");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(5000);
            sheet.Cells["C2"].PutValue(3000);

            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(7000);
            sheet.Cells["C3"].PutValue(4000);

            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(6000);
            sheet.Cells["C4"].PutValue(3500);

            // Add a pivot table based on the data range A1:C4
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");      // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");    // Data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Cost");       // Data field

            // Add a calculated field named "Profit" with the formula [Revenue]-[Cost]
            // The formula must start with '=' when using AddCalculatedField
            pivotTable.AddCalculatedField("Profit", "=Revenue-Cost", true);

            // Refresh and calculate the pivot table to populate the new field
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_With_Profit_CalculatedField.xlsx");
        }
    }
}