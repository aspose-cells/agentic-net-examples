using System;
using Aspose.Cells;

class ConvertWorkbookToCsv
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(sourcePath);

        // Remove all comments from every worksheet to keep the CSV output clean
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments();
        }

        // Prepare CSV save options (export all sheets, not just the active one)
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.ExportAllSheets = true;   // optional: set to false to export only the active sheet

        // Save the workbook as CSV using the save options
        workbook.Save("output.csv", csvOptions);
    }
}