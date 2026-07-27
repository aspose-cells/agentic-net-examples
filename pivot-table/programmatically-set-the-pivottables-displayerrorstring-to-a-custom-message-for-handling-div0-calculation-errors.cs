using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotErrorStringDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source data
            // Column A: Category, Column B: Value, Column C: Formula that may cause #DIV/0!
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["C1"].PutValue("Formula");

            cells["A2"].PutValue("Item1");
            cells["B2"].PutValue(100);
            cells["C2"].Formula = "=B2/0"; // Will generate #DIV/0!

            cells["A3"].PutValue("Item2");
            cells["B3"].PutValue(200);
            cells["C3"].Formula = "=B3/2"; // Normal calculation

            // Create a pivot table based on the source range (including the formula column)
            int pivotIndex = sheet.PivotTables.Add("A1:C3", "E5", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);  // Formula column (will contain errors)

            // Enable custom error string display and set the custom message
            pivotTable.DisplayErrorString = true;
            pivotTable.ErrorString = "Division error";

            // Refresh data and calculate the pivot table to apply settings
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableErrorStringDemo.xlsx");

            Console.WriteLine("PivotTable created with custom error string for #DIV/0! errors.");
        }
    }
}