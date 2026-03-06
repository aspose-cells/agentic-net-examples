using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create load options for XLSX format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        // Enable parsing of pivot cached records (equivalent to LoadPivotTableCache = true)
        loadOptions.ParsingPivotCachedRecords = true;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Save the workbook (the pivot cache will be loaded as per the option)
        workbook.Save("output.xlsx");
    }
}