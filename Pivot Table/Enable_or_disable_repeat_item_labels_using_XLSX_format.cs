using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRepeatItemLabels
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a pivot table
            Workbook workbook = new Workbook("input.xlsx");

            // Access the worksheet that holds the pivot table (by name or index)
            Worksheet worksheet = workbook.Worksheets["Sheet1"]; // adjust name as needed

            // Get the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Enable repeating item labels for the first row field
            if (pivotTable.RowFields.Count > 0)
            {
                pivotTable.RowFields[0].IsRepeatItemLabels = true;
            }

            // Enable repeating item labels for the first column field
            if (pivotTable.ColumnFields.Count > 0)
            {
                pivotTable.ColumnFields[0].IsRepeatItemLabels = true;
            }

            // Example: disable repeating item labels for the second row field (if exists)
            if (pivotTable.RowFields.Count > 1)
            {
                pivotTable.RowFields[1].IsRepeatItemLabels = false;
            }

            // Save the modified workbook in XLSX format
            workbook.Save("output.xlsx");
        }
    }
}