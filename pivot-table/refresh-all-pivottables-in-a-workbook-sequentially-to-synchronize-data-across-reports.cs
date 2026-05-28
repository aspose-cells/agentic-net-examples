using System;
using Aspose.Cells;

class RefreshAllPivotTables
{
    static void Main()
    {
        // Load the workbook that contains the pivot tables
        Workbook workbook = new Workbook("input.xlsx");

        // Refresh every pivot table in every worksheet of the workbook
        workbook.Worksheets.RefreshPivotTables();

        // Save the workbook after the refresh operation
        workbook.Save("output.xlsx");
    }
}