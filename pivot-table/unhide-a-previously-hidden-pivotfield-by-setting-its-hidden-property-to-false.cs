using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotUnhideField
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a pivot table
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Identify the pivot field to unhide.
            // Example: unhide the first row field. Change the index or use a name as required.
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("No row fields found in the pivot table.");
                return;
            }

            PivotField pivotField = pivotTable.RowFields[0]; // or pivotTable.RowFields["FieldName"]

            // Unhide all items in the selected pivot field
            for (int i = 0; i < pivotField.ItemCount; i++)
            {
                // Set hidden flag to false for each item
                pivotField.HideItem(i, false);
            }

            // Refresh and recalculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}