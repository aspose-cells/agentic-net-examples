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
            // Column A – Category, Column B – Value
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Value";
            cells["A2"].Value = "B";
            cells["A3"].Value = "A";
            cells["A4"].Value = "C";
            cells["A5"].Value = "D";
            cells["B2"].Value = 40;
            cells["B3"].Value = 10;
            cells["B4"].Value = 30;
            cells["B5"].Value = 20;

            // Add a pivot table based on the data range A1:B5
            int ptIndex = sheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Add the Category field as a row field and the Value field as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // ------------------------------------------------------------
            // Apply an ascending custom sort order to the row field
            // ------------------------------------------------------------
            // Get the row field (Category)
            PivotField rowField = pivotTable.RowFields[0];

            // Enable automatic sorting
            rowField.IsAutoSort = true;

            // Set the sort direction to ascending
            rowField.IsAscendSort = true;

            // AutoSortField = -1 means the field is sorted by its own labels
            // (i.e., the Category values themselves)
            rowField.AutoSortField = -1;

            // Optional: let the pivot table respect built‑in custom list sorting
            // (e.g., a user‑defined order like A, B, C, D). This does not create
            // a custom list but enables the feature if such a list exists.
            pivotTable.CustomListSort = true;

            // Refresh the pivot cache and calculate the pivot table
            pivotTable.RefreshDataFlag = true;
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_CustomAscendingSort.xlsx");
        }
    }
}