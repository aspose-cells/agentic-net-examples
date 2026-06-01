using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCalculatedFieldDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            // Header row
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Sales");

            // Data rows
            cells["A2"].PutValue("North");
            cells["B2"].PutValue(2500);
            cells["A3"].PutValue("South");
            cells["B3"].PutValue(1500);
            cells["A4"].PutValue("East");
            cells["B4"].PutValue(800);
            cells["A5"].PutValue("West");
            cells["B5"].PutValue(3000);
            cells["A6"].PutValue("North");
            cells["B6"].PutValue(1200);
            cells["A7"].PutValue("South");
            cells["B7"].PutValue(500);

            // Add a pivot table based on the data range A1:B7, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:B7", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");   // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");   // Data field

            // Add a calculated field that categorizes sales:
            // High  : Sales > 2000
            // Medium: Sales > 1000
            // Low   : otherwise
            string formula = "=IF(Sales>2000,\"High\",IF(Sales>1000,\"Medium\",\"Low\"))";
            pivotTable.AddCalculatedField("SalesCategory", formula, true);

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_With_CalculatedField.xlsx");
        }
    }
}