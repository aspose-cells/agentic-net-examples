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
            // Column A: Product (row field)
            // Column B: Sales (numeric data field)
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "B";
            cells["A3"].Value = "A";
            cells["A4"].Value = "C";
            cells["B2"].Value = 200;
            cells["B3"].Value = 300;
            cells["B4"].Value = 100;

            // Add a pivot table covering the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the row field (Product) and the data field (Sales)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the row field object
            PivotField rowField = pivotTable.RowFields[0];

            // Enable automatic sorting and set it to ascending
            rowField.IsAutoSort = true;          // Turn on auto‑sorting
            rowField.IsAscendSort = true;        // Ascending order

            // Specify that sorting should be based on the first data field (Sales)
            // Index 0 refers to the first data field added to the pivot table
            rowField.AutoSortField = 0;

            // Refresh the pivot table to apply sorting
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_SortedByNumericValues.xlsx");
        }
    }
}