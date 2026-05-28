using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class RefreshPivotAfterLoad
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Create load options and enable parsing of pivot cached records
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.ParsingPivotCachedRecords = true;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Refresh all pivot tables in the workbook
        workbook.Worksheets.RefreshPivotTables();

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}