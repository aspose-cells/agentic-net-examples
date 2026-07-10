using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Prepare load options and disable parsing of pivot cached records
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingPivotCachedRecords = false;   // default is false, set explicitly as required

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Refresh all pivot tables in the workbook (ensures they reflect the current source data)
        workbook.Worksheets.RefreshPivotTables();

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}