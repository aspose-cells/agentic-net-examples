using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class AddCountDataFieldDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample source data
        cells["A1"].Value = "Item";
        cells["B1"].Value = "Value";
        cells["A2"].Value = "A";
        cells["B2"].Value = 10;
        cells["A3"].Value = "B";
        cells["B3"].Value = 20;
        cells["A4"].Value = "A";
        cells["B4"].Value = 30;
        cells["A5"].Value = "C";
        cells["B5"].Value = 40;

        // Add a pivot table based on the source range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the row field (Item)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Item");

        // Add the data field (Value) and set its aggregation to Count
        int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");
        PivotField dataField = pivotTable.DataFields[dataFieldPos];
        dataField.Function = ConsolidationFunction.Count;

        // Refresh and calculate the pivot table
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotTable_With_CountField.xlsx");
    }
}