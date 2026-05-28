using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotRowCountLogger
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("A");
            sheet.Cells["B5"].PutValue(40);
            sheet.Cells["A6"].PutValue("B");
            sheet.Cells["B6"].PutValue(50);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "MyPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row field, Value as data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh the pivot table to ensure it reflects the source data
            sheet.RefreshPivotTables();

            // After refresh, obtain the number of distinct row items generated
            // This can be retrieved from the first row field's PivotItems collection
            int rowItemCount = pivot.RowFields[0].PivotItems.Count;

            // Log the row count for diagnostics
            Console.WriteLine($"PivotTable '{pivot.Name}' generated {rowItemCount} row items after refresh.");

            // Save the workbook (optional, demonstrates full lifecycle)
            workbook.Save("PivotRowCountDemo.xlsx");
        }
    }
}