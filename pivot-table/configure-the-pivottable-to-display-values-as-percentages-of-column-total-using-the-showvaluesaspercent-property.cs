using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    class ShowValuesAsPercentageOfColumn
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            // Columns: Product | Quarter | Sales
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Quarter";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = "Q1";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Apple";
            sheet.Cells["B3"].Value = "Q2";
            sheet.Cells["C3"].Value = 1500;

            sheet.Cells["A4"].Value = "Orange";
            sheet.Cells["B4"].Value = "Q1";
            sheet.Cells["C4"].Value = 800;

            sheet.Cells["A5"].Value = "Orange";
            sheet.Cells["B5"].Value = "Q2";
            sheet.Cells["C5"].Value = 950;

            // Add a pivot table based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivot = pivots[pivotIndex];

            // Add fields to the pivot table
            // Row field: Product
            int rowFieldIdx = pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            // Column field: Quarter
            int columnFieldIdx = pivot.AddFieldToArea(PivotFieldType.Column, "Quarter");
            // Data field: Sales
            int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the data field object
            PivotField dataField = pivot.DataFields[dataFieldIdx];

            // Configure the data field to show values as percentage of column total
            // Using ShowValuesAs method with PercentageOfColumn display format.
            // Base field is the column field (Quarter), base item position type can be Next (default) and base item index is 0.
            dataField.ShowValuesAs(
                PivotFieldDataDisplayFormat.PercentageOfColumn,
                columnFieldIdx,               // base field index (Quarter)
                PivotItemPositionType.Next,   // base item position type
                0);                           // base item index (ignored for Next)

            // Refresh and calculate the pivot table to apply the settings
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("Pivot_ShowValuesAs_PercentageOfColumn.xlsx");
        }
    }
}