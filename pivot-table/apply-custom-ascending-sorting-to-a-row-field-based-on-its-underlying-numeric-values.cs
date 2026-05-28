using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSortingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            // Header row
            cells["A1"].Value = "Product";
            cells["B1"].Value = "UnitsSold";

            // Data rows (numeric values in column B)
            cells["A2"].Value = "Widget";
            cells["B2"].Value = 150;
            cells["A3"].Value = "Gadget";
            cells["B3"].Value = 45;
            cells["A4"].Value = "Doohickey";
            cells["B4"].Value = 300;
            cells["A5"].Value = "Thingamajig";
            cells["B5"].Value = 120;

            // Add a pivot table based on the data range A1:B5, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the product column as a row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Add the units sold column as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "UnitsSold");

            // Retrieve the row field (the first row field)
            PivotField rowField = pivotTable.RowFields[0];

            // Enable automatic sorting and set it to ascending
            rowField.IsAutoSort = true;          // Turn on auto‑sorting
            rowField.IsAscendSort = true;        // Ascending order

            // Specify that sorting should be based on the first data field (UnitsSold)
            // Index 0 refers to the first data field added to the pivot table
            rowField.AutoSortField = 0;

            // Alternatively, you can achieve the same with the SortBy method:
            // rowField.SortBy(SortOrder.Ascending, 0);

            // Refresh the pivot table to apply sorting
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_CustomAscendingSort.xlsx");
        }
    }
}