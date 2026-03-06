using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

public class SetPivotTableNameDemo
{
    public static void Run()
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
        sheet.Cells["B4"].PutValue(20);

        // Add a pivot table; initial name is "PivotTable1"
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Set a custom name for the pivot table
        pivotTable.Name = "MyCustomPivot";

        // Configure the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Quantity as data field
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("PivotTableWithCustomName.xlsx", SaveFormat.Xlsx);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        SetPivotTableNameDemo.Run();
    }
}