using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableRefreshDateDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Fruit");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(15);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(8);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "C3", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table (optional, but typical)
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Quantity column as data field

            // Refresh the pivot table to set the RefreshDate property
            pivot.RefreshData();
            pivot.CalculateData();

            // Output the refresh date
            Console.WriteLine("Pivot Table Refresh Date (initial): " + pivot.RefreshDate.ToLongDateString());

            // Save the workbook to an XLSX file
            string filePath = "PivotRefreshDateDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Load the workbook back to verify that the refresh date persists
            Workbook loadedWb = new Workbook(filePath);
            PivotTable loadedPivot = loadedWb.Worksheets[0].PivotTables[0];

            // Output the refresh date after reloading
            Console.WriteLine("Pivot Table Refresh Date (after reload): " + loadedPivot.RefreshDate.ToLongDateString());
        }
    }
}