using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class ApplyCompactLayout
{
    static void Main()
    {
        // Load an existing workbook that already contains a pivot table
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Verify that the worksheet has at least one pivot table
        if (sheet.PivotTables.Count > 0)
        {
            // Retrieve the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Apply the compact layout to the pivot table
            pivotTable.ShowInCompactForm();

            // Refresh the data source and recalculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();
        }

        // Save the workbook with the updated layout
        workbook.Save("output.xlsx");
    }
}