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

            // Populate sample data with Revenue and Cost columns
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Revenue";
            cells["C1"].Value = "Cost";

            cells["A2"].Value = "A";
            cells["B2"].Value = 1200;
            cells["C2"].Value = 800;

            cells["A3"].Value = "B";
            cells["B3"].Value = 1500;
            cells["C3"].Value = 900;

            cells["A4"].Value = "C";
            cells["B4"].Value = 1800;
            cells["C4"].Value = 1100;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Cost");

            // Add a calculated field named "Profit" with the formula [Revenue]-[Cost]
            // The formula must start with '=' and use the field names as they appear in the source data
            pivotTable.AddCalculatedField("Profit", "=Revenue-Cost");

            // Refresh and calculate the pivot table to populate the new field
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_With_Profit_CalculatedField.xlsx");
        }
    }
}