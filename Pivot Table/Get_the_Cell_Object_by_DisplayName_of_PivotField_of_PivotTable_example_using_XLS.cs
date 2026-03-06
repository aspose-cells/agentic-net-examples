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
        cells["B3"].Value = 15;
        cells["A4"].Value = "Banana";
        cells["B4"].Value = 20;

        // Add a pivot table to the worksheet
        int ptIndex = sheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];

        // Add row and data fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

        // Refresh data and calculate the pivot table
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Get the display name of the first data field
        string displayName = pivotTable.DataFields[0].DisplayName;

        // Retrieve the cell that corresponds to the display name
        Cell cell = pivotTable.GetCellByDisplayName(displayName);

        // Output information about the retrieved cell
        Console.WriteLine($"Display Name: {displayName}");
        Console.WriteLine($"Cell Name: {(cell != null ? cell.Name : "null")}");
        Console.WriteLine($"Cell Value: {(cell != null ? cell.Value?.ToString() ?? "null" : "null")}");

        // Save the workbook in XLSX format
        workbook.Save("PivotTable_GetCellByDisplayName_Output.xlsx");
    }
}