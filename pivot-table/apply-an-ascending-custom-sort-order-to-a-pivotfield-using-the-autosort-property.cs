using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCustomSort
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "B";
            cells["A3"].Value = "A";
            cells["A4"].Value = "C";
            cells["B2"].Value = 200;
            cells["B3"].Value = 100;
            cells["B4"].Value = 300;

            // Add a pivot table based on the data range
            int ptIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Add the "Category" field as a row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the "Amount" field as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Retrieve the row field we just added
            PivotField rowField = pivotTable.RowFields[0];

            // Enable automatic sorting and set it to ascending order
            rowField.IsAutoSort = true;          // Turn on auto‑sort
            rowField.IsAscendSort = true;        // Ascending sort
            rowField.AutoSortField = -1;         // Sort by the field's own labels

            // Refresh the pivot table data and calculate the results
            pivotTable.RefreshDataFlag = true;
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_CustomAscendingSort.xlsx");
        }
    }
}