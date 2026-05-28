using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotErrorStringDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data that will cause a division by zero error in a calculated column
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["C1"].PutValue("DivByZero"); // This column will hold a formula =B1/0

            cells["A2"].PutValue("Item1");
            cells["B2"].PutValue(100);
            cells["C2"].Formula = "=B2/0"; // Generates #DIV/0! error

            cells["A3"].PutValue("Item2");
            cells["B3"].PutValue(200);
            cells["C3"].Formula = "=B3/0"; // Generates #DIV/0! error

            // Create a pivot table based on the data range (including the error column)
            int pivotIndex = sheet.PivotTables.Add("A1:C3", "E5", "ErrorPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);  // DivByZero (error column)

            // Enable custom error string display and set the custom message
            pivotTable.DisplayErrorString = true;
            pivotTable.ErrorString = "Division by zero not allowed";

            // Refresh data and calculate the pivot table to apply settings
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableWithCustomErrorString.xlsx");

            Console.WriteLine("PivotTable created with custom error string for #DIV/0! errors.");
        }
    }
}