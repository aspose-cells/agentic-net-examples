using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCalculatedFieldDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Product, Revenue, Profit
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Revenue");
            cells["C1"].PutValue("Profit");

            cells["A2"].PutValue("A");
            cells["B2"].PutValue(1000);
            cells["C2"].PutValue(200);

            cells["A3"].PutValue("B");
            cells["B3"].PutValue(1500);
            cells["C3"].PutValue(300);

            cells["A4"].PutValue("C");
            cells["B4"].PutValue(2000);
            cells["C4"].PutValue(500);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");          // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");        // Data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Profit");         // Data field

            // Add a calculated field "ProfitMargin" = Profit / Revenue and drag it to the data area
            pivotTable.AddCalculatedField("ProfitMargin", "=Profit/Revenue", true);

            // Optionally format the calculated field as percentage
            PivotField profitMarginField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];
            profitMarginField.NumberFormat = "0.00%";

            // Refresh and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_With_ProfitMargin.xlsx");
        }
    }
}