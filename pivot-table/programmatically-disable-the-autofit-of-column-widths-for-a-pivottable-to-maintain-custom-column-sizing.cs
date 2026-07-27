using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load an existing workbook that contains a pivot table
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet where the pivot table resides (adjust name or index as needed)
        Worksheet worksheet = workbook.Worksheets["Sheet1"]; // or workbook.Worksheets[0];

        // Ensure the worksheet has at least one pivot table
        if (worksheet.PivotTables.Count > 0)
        {
            // Get the first pivot table in the collection
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Disable automatic column width adjustment when the pivot table is refreshed or updated
            pivotTable.AutofitColumnWidthOnUpdate = false;

            // Refresh pivot tables to apply the change
            worksheet.RefreshPivotTables();
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}