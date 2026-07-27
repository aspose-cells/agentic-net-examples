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

        // Fill sample data for the pivot table
        Cells cells = sheet.Cells;
        cells["A1"].Value = "Fruit";
        cells["B1"].Value = "Quantity";
        cells["A2"].Value = "Apple";
        cells["B2"].Value = 10;
        cells["A3"].Value = "Orange";
        cells["B3"].Value = 15;
        cells["A4"].Value = "Banana";
        cells["B4"].Value = 20;

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "C3", "FruitPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure the pivot table: Fruit as rows, Quantity as data
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

        // Optional: do not embed the source data with the file
        pivot.SaveData = false;

        // Save the workbook to XLSX format using default save options
        workbook.Save("PivotTableDemo.xlsx");
    }
}