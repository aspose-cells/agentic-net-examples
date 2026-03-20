using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Disable keeping unparsed data to improve performance
        loadOptions.KeepUnparsedData = false;

        // Load only cell values (no formatting, formulas, etc.)
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellValue);

        // Open the workbook with the configured load options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the first worksheet and display a cell value
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("A1 value: " + sheet.Cells["A1"].Value);
    }
}