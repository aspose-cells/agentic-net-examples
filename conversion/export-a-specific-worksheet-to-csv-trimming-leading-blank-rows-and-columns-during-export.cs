using System;
using Aspose.Cells;

class ExportWorksheetToCsv
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Choose the worksheet you want to export by setting it as the active sheet
        // For example, export the second worksheet (index starts at 0)
        workbook.Worksheets.ActiveSheetIndex = 1;

        // Configure CSV (text) save options
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);

        // Ensure leading blank rows and columns are trimmed (default is true, set explicitly for clarity)
        saveOptions.TrimLeadingBlankRowAndColumn = true;

        // Export only the active worksheet (default ExportAllSheets is false, set explicitly)
        saveOptions.ExportAllSheets = false;

        // Save the selected worksheet to a CSV file
        workbook.Save("output.csv", saveOptions);
    }
}