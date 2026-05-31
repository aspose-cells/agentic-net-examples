using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class PivotShowValuesAsPercentageOfColumn
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Product";
        sheet.Cells["B1"].Value = "Quarter";
        sheet.Cells["C1"].Value = "Sales";

        sheet.Cells["A2"].Value = "Apple";
        sheet.Cells["B2"].Value = "Q1";
        sheet.Cells["C2"].Value = 1000;

        sheet.Cells["A3"].Value = "Apple";
        sheet.Cells["B3"].Value = "Q2";
        sheet.Cells["C3"].Value = 1500;

        sheet.Cells["A4"].Value = "Orange";
        sheet.Cells["B4"].Value = "Q1";
        sheet.Cells["C4"].Value = 800;

        sheet.Cells["A5"].Value = "Orange";
        sheet.Cells["B5"].Value = "Q2";
        sheet.Cells["C5"].Value = 1200;

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("=Sheet1!A1:C5", "E3", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add fields: Product as rows, Quarter as columns, Sales as data
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");
        int columnFieldIdx = pivot.AddFieldToArea(PivotFieldType.Column, "Quarter");
        int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Retrieve the column field (base field) and the data field
        PivotField columnField = pivot.ColumnFields[columnFieldIdx];
        PivotField dataField = pivot.DataFields[dataFieldIdx];

        // Configure the data field to show values as percentage of column total
        dataField.ShowValuesAs(
            PivotFieldDataDisplayFormat.PercentageOfColumn, // display format
            columnField.BaseIndex,                         // base field index
            PivotItemPositionType.Next,                    // base item position type (not used for this format)
            0);                                            // base item index (not used for this format)

        // Refresh and calculate the pivot table to apply the settings
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook with the configured pivot table
        workbook.Save("PivotShowValuesAsPercentageOfColumn.xlsx");
    }
}