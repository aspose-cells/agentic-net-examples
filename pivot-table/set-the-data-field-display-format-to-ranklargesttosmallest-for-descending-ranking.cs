using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].Value = "Fruit";
        cells["B1"].Value = "Quantity";
        cells["A2"].Value = "Apple";
        cells["B2"].Value = 10;
        cells["A3"].Value = "Orange";
        cells["B3"].Value = 20;
        cells["A4"].Value = "Banana";
        cells["B4"].Value = 15;
        cells["A5"].Value = "Apple";
        cells["B5"].Value = 5;

        // Add a pivot table covering the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "C3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add a row field (Fruit) to the pivot table
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");

        // Add a data field (Quantity) to the pivot table
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
        PivotField dataField = pivot.DataFields[0];

        // Set the data field display format to RankLargestToSmallest (descending ranking)
        dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.RankLargestToSmallest;

        // Calculate the pivot data and save the workbook
        pivot.CalculateData();
        workbook.Save("PivotRankLargestToSmallest.xlsx");
    }
}