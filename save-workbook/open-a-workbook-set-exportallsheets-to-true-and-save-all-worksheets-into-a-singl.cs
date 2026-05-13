using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create save options for CSV format
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        // Export all worksheets into the same CSV file
        saveOptions.ExportAllSheets = true;

        // Save the workbook; all sheets will be concatenated into a single CSV file
        workbook.Save("all_sheets.csv", saveOptions);
    }
}