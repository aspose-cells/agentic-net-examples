using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCountExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample source data
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Item";
            cells["A2"].Value = "A";
            cells["B2"].Value = "Item1";
            cells["A3"].Value = "A";
            cells["B3"].Value = "Item2";
            cells["A4"].Value = "B";
            cells["B4"].Value = "Item3";
            cells["A5"].Value = "B";
            cells["B5"].Value = "Item4";

            // Add a pivot table based on the source range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the row field (Category)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add a new data field (Item) and set its aggregation function to Count
            int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, "Item");
            PivotField dataField = pivotTable.DataFields[dataFieldPos];
            dataField.Function = ConsolidationFunction.Count;

            // Refresh and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableCountField.xlsx");
        }
    }
}