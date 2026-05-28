using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCustomNumericGrouping
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Sales amounts
            sheet.Cells["A1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue(500);
            sheet.Cells["A3"].PutValue(1500);
            sheet.Cells["A4"].PutValue(3200);
            sheet.Cells["A5"].PutValue(4700);
            sheet.Cells["A6"].PutValue(8200);
            sheet.Cells["A7"].PutValue(11500);

            // Add a pivot table based on the data range
            // Source range: A1:A7, Destination: C3, Name: "SalesPivot"
            int pivotIndex = sheet.PivotTables.Add("A1:A7", "C3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Sales field to the row area of the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);

            // Retrieve the pivot field representing the Sales column
            PivotField salesField = pivotTable.RowFields[0];

            // Define custom numeric grouping:
            //   Start value : 0
            //   End value   : 12000
            //   Interval    : 2500 (creates ranges 0‑2500, 2500‑5000, 5000‑7500, 7500‑10000, 10000‑12000)
            //   newField    : false (grouping applied to the existing field)
            salesField.GroupBy(0, 12000, 2500, false);

            // Save the workbook with the grouped pivot table
            workbook.Save("CustomNumericGrouping.xlsx");
        }
    }
}