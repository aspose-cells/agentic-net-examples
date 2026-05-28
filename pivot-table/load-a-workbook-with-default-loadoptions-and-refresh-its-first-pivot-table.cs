using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class RefreshFirstPivotTable
{
    static void Main()
    {
        // Path to the workbook to be loaded
        string inputPath = "input.xlsx";

        // Create default load options
        LoadOptions loadOptions = new LoadOptions();

        // Load the workbook with the default load options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Access the first worksheet (assumed to contain the pivot table)
        Worksheet sheet = workbook.Worksheets[0];

        // Refresh the first pivot table if it exists
        if (sheet.PivotTables.Count > 0)
        {
            PivotTable pivot = sheet.PivotTables[0];
            pivot.RefreshData();      // Refresh data from the source
            pivot.CalculateData();    // Recalculate the pivot table values
        }

        // Save the workbook after refreshing the pivot table
        workbook.Save("output.xlsx");
    }
}