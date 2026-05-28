using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class RetrievePivotTableRefreshedByWho
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Fruit");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["B3"].PutValue(15);
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B4"].PutValue(8);

        // Add a pivot table to the worksheet
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "FruitPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Configure the pivot table (row field and data field)
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Quantity column

        // Refresh the pivot table to populate RefreshDate and RefreshedByWho
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Retrieve and display the user who last refreshed the pivot table
        Console.WriteLine("Last Refresh Date: " + pivotTable.RefreshDate);
        Console.WriteLine("Refreshed By: " + pivotTable.RefreshedByWho);

        // Save the workbook (optional, demonstrates persistence)
        string outputPath = "PivotRefreshedByWho.xlsx";
        workbook.Save(outputPath);
    }
}