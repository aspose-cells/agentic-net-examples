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

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Product";
        sheet.Cells["B1"].Value = "Sales";
        sheet.Cells["A2"].Value = "B";
        sheet.Cells["A3"].Value = "C";
        sheet.Cells["A4"].Value = "A";
        sheet.Cells["B2"].Value = 200;
        sheet.Cells["B3"].Value = 300;
        sheet.Cells["B4"].Value = 100;

        // Add a pivot table covering the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add a row field (Product) and a data field (Sales)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Retrieve the row field to configure auto‑sorting
        PivotField rowField = pivotTable.RowFields[0];

        // Enable auto sort, set descending order, and sort by the first data field (index 0)
        rowField.IsAutoSort = true;
        rowField.IsAscendSort = false; // false = descending
        rowField.AutoSortField = 0;    // sort by the first data field (Sales)

        // Refresh the pivot table data and calculate the results
        pivotTable.RefreshDataFlag = true;
        pivotTable.CalculateData();

        // Save the workbook with the applied sorting
        workbook.Save("PivotFieldDescendingAutoSort.xlsx");
    }
}