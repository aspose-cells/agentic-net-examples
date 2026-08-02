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
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["A2"].Value = "A";
        sheet.Cells["A3"].Value = "B";
        sheet.Cells["A4"].Value = "A";
        sheet.Cells["A5"].Value = "B";

        sheet.Cells["B1"].Value = "Amount";
        sheet.Cells["B2"].Value = 100;
        sheet.Cells["B3"].Value = 200;
        sheet.Cells["B4"].Value = 150;
        sheet.Cells["B5"].Value = 250;

        // Add a pivot table covering the data range and place it at D1
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the Category field to the row area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

        // Add the Amount field to the data area
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Retrieve the row field (Category) to which we will add a calculated item
        PivotField categoryField = pivotTable.RowFields[0];

        // Add a calculated item that sums the values of categories A and B
        // The formula uses the item names as they appear in the source data
        categoryField.AddCalculatedItem("Total_AB", "=A + B");

        // Refresh the pivot table data and recalculate to reflect the new item
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the calculated item added
        workbook.Save("CalculatedItemDemo.xlsx");
    }
}