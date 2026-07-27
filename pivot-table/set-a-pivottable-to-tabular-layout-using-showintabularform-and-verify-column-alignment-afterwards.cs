using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("SubCategory");
        dataSheet.Cells["C1"].PutValue("Amount");
        dataSheet.Cells["A2"].PutValue("Fruit");
        dataSheet.Cells["B2"].PutValue("Apple");
        dataSheet.Cells["C2"].PutValue(100);
        dataSheet.Cells["A3"].PutValue("Fruit");
        dataSheet.Cells["B3"].PutValue("Banana");
        dataSheet.Cells["C3"].PutValue(150);
        dataSheet.Cells["A4"].PutValue("Vegetable");
        dataSheet.Cells["B4"].PutValue("Carrot");
        dataSheet.Cells["C4"].PutValue(80);

        // Add a worksheet for the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create the pivot table
        int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C4", "A1", "PivotTable1");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Set the pivot table to Tabular layout
        pivotTable.ShowInTabularForm();

        // Refresh and calculate the pivot data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Verify column alignment:
        // In Tabular form each field occupies its own column, so the start column of the
        // column header range should match the start column of the data body range.
        CellArea columnRange = pivotTable.ColumnRange;
        CellArea dataBodyRange = pivotTable.DataBodyRange;

        bool isAligned = columnRange.StartColumn == dataBodyRange.StartColumn;

        Console.WriteLine($"Column alignment verification: {(isAligned ? "PASS" : "FAIL")}");
        Console.WriteLine($"ColumnRange.StartColumn = {columnRange.StartColumn}, DataBodyRange.StartColumn = {dataBodyRange.StartColumn}");

        // Save the workbook
        workbook.Save("PivotTabularLayoutDemo.xlsx");
    }
}