using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshedByDemo
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
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table (row field and data field)
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Quantity column

            // Refresh the pivot table to set the RefreshedByWho property
            pivot.RefreshData();
            pivot.CalculateData();

            // Output the name of the user who refreshed the pivot table
            Console.WriteLine("Refreshed By (in-memory): " + pivot.RefreshedByWho);
            Console.WriteLine("Refresh Date (in-memory): " + pivot.RefreshDate);

            // Save the workbook to an XLSX file
            string filePath = "PivotRefreshedByDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Load the workbook from the saved file
            Workbook loadedWb = new Workbook(filePath);
            PivotTable loadedPivot = loadedWb.Worksheets[0].PivotTables[0];

            // Access the RefreshedByWho property after reload
            Console.WriteLine("Refreshed By (after reload): " + loadedPivot.RefreshedByWho);
            Console.WriteLine("Refresh Date (after reload): " + loadedPivot.RefreshDate);
        }
    }
}