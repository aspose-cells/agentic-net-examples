using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class DisableRowGrandTotals
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Product";
        sheet.Cells["B1"].Value = "Region";
        sheet.Cells["C1"].Value = "Sales";

        sheet.Cells["A2"].Value = "Product A";
        sheet.Cells["B2"].Value = "North";
        sheet.Cells["C2"].Value = 1000;

        sheet.Cells["A3"].Value = "Product B";
        sheet.Cells["B3"].Value = "South";
        sheet.Cells["C3"].Value = 2000;

        sheet.Cells["A4"].Value = "Product A";
        sheet.Cells["B4"].Value = "West";
        sheet.Cells["C4"].Value = 1500;

        // Add a pivot table based on the data range
        PivotTableCollection pivotTables = sheet.PivotTables;
        int pivotIndex = pivotTables.Add("A1:C4", "E3", "PivotTable1");
        PivotTable pivotTable = pivotTables[pivotIndex];

        // Configure pivot fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Disable row grand totals
        pivotTable.ShowRowGrandTotals = false;

        // Calculate the pivot table data
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotTable_NoRowGrandTotals.xlsx");
    }
}