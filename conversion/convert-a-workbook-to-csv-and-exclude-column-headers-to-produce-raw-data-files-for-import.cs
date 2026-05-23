using System;
using Aspose.Cells;

class ConvertWorkbookToCsvWithoutHeaders
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the first row of the first worksheet contains column headers.
        // Delete that row so it will not appear in the exported CSV.
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.DeleteRow(0);

        // Configure CSV save options.
        // ExportAllSheets = true will include every worksheet; set to false to export only the active sheet.
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.ExportAllSheets = true;
        // Keep leading blank rows/columns as they are (optional, based on requirements).
        csvOptions.TrimLeadingBlankRowAndColumn = false;

        // Save the workbook as a CSV file without the header row.
        workbook.Save("output.csv", csvOptions);
    }
}