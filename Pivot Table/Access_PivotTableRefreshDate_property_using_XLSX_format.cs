using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDateDemo
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
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "FruitPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table (row field and data field)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Quantity column

            // Refresh the pivot table to set the RefreshDate property
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Display the refresh date before saving
            Console.WriteLine("Refresh Date before saving: " + pivotTable.RefreshDate.ToLongDateString());

            // Save the workbook in XLSX format
            string filePath = "PivotRefreshDateDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Load the workbook back
            Workbook loadedWorkbook = new Workbook(filePath);
            PivotTable loadedPivot = loadedWorkbook.Worksheets[0].PivotTables[0];

            // Display the refresh date after reloading
            Console.WriteLine("Refresh Date after reloading: " + loadedPivot.RefreshDate.ToLongDateString());
        }
    }
}