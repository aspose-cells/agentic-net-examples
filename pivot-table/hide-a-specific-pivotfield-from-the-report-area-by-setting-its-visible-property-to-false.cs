using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotFieldHideDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["A2"].PutValue("PO-23-05");
            sheet.Cells["A3"].PutValue("PO-23-06");
            sheet.Cells["A4"].PutValue("PO-23-05");
            sheet.Cells["A5"].PutValue("PO-23-07");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["B4"].PutValue(1500);
            sheet.Cells["B5"].PutValue(3000);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Product" field to the Row area and "Sales" to the Data area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Hide the "Product" field from the report area.
            // Since PivotField does not expose a Visible property, we remove it from its area.
            // This effectively hides the field from the pivot table view.
            pivotTable.RemoveField(PivotFieldType.Row, "Product");

            // Refresh and calculate the pivot table after the modification
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotFieldHiddenDemo.xlsx");
        }
    }
}