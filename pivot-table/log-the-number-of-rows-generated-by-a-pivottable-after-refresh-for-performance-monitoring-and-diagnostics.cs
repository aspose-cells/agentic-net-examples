using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotRowCountLogger
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("C");
            sheet.Cells["B5"].PutValue(300);

            // Add a pivot table based on the source data
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "MyPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh the pivot table to ensure it reflects the current data
            // (lifecycle: refresh)
            sheet.RefreshPivotTables();

            // After refresh, obtain the number of distinct row items generated
            // This corresponds to the count of pivot items in the first row field
            int rowItemCount = pivot.RowFields[0].PivotItems.Count;

            // Log the row count for diagnostics
            Console.WriteLine($"PivotTable row item count after refresh: {rowItemCount}");

            // Save the workbook (lifecycle: save)
            workbook.Save("PivotRowCountLog.xlsx");
        }
    }
}