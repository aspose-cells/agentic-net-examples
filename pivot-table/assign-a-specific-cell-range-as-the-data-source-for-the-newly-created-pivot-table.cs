using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (initial source)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(300);

            // Add a pivot table with an initial source range (will be changed later)
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "MyPivotTable");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table (using the initial source)
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Define a new data source range (e.g., C1:D4 on the same sheet)
            // First, add new data for the new source
            sheet.Cells["C1"].PutValue("Category");
            sheet.Cells["D1"].PutValue("Amount");
            sheet.Cells["C2"].PutValue("X");
            sheet.Cells["D2"].PutValue(400);
            sheet.Cells["C3"].PutValue("Y");
            sheet.Cells["D3"].PutValue(500);
            sheet.Cells["C4"].PutValue("Z");
            sheet.Cells["D4"].PutValue(600);

            // Change the pivot table's data source to the new range
            string[] newSource = new string[] { "C1:D4" };
            pivot.ChangeDataSource(newSource);

            // Refresh and recalculate the pivot table to reflect the new source
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableWithNewDataSource.xlsx");
        }
    }
}