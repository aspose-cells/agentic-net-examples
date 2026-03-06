using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class LoadPivotCacheDemo
{
    static void Main()
    {
        // Create load options and enable parsing of pivot cached records
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingPivotCachedRecords = true;

        // Load an existing XLSX workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // If the worksheet contains pivot tables, work with the first one
        if (worksheet.PivotTables.Count > 0)
        {
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Example operation: ensure the pivot data is refreshed (uses cached records)
            pivotTable.RefreshDataFlag = true;

            // Output some basic information about the pivot table
            Console.WriteLine("Pivot Table Name: " + pivotTable.Name);
            Console.WriteLine("Row Fields Count: " + pivotTable.RowFields.Count);
            Console.WriteLine("Data Fields Count: " + pivotTable.DataFields.Count);
        }

        // Save the workbook (optional, demonstrates that the workbook can still be saved)
        workbook.Save("output.xlsx");
    }
}