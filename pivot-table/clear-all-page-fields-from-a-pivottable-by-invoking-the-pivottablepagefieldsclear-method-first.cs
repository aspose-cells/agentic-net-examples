using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ClearPivotTablePageFieldsDemo
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("P1");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(1000);

            sheet.Cells["A3"].PutValue("P2");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(1500);

            sheet.Cells["A4"].PutValue("P1");
            sheet.Cells["B4"].PutValue("East");
            sheet.Cells["C4"].PutValue(2000);

            // Add a pivot table with a page field (Region)
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Region");

            // Refresh and calculate to build the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Clear all page fields from the pivot table
            pivotTable.PageFields.Clear();

            // Recalculate after clearing page fields (optional)
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("ClearPivotTablePageFieldsDemo.xlsx");
        }
    }
}