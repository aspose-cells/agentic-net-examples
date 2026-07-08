using System;
using System.Data;
using Aspose.Cells;

class PreserveHtmlImportDemo
{
    static void Main()
    {
        // Create a DataTable containing HTML strings
        DataTable table = new DataTable();
        table.Columns.Add("ID");
        table.Columns.Add("Content");
        table.Rows.Add("1", "<b>Bold Text</b>");
        table.Rows.Add("2", "<a href='https://www.example.com'>Example Link</a>");

        // Create a new workbook and get the first worksheet's cells
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Configure import options to preserve HTML formatting
        ImportTableOptions importOptions = new ImportTableOptions
        {
            IsFieldNameShown = true, // import column headers
            IsHtmlString = true      // treat cell values as HTML
        };

        // Import the DataTable starting at cell A1 (row 0, column 0)
        cells.ImportData(table, 0, 0, importOptions);

        // Save the workbook to an Excel file
        workbook.Save("PreserveHtmlDemo.xlsx");
    }
}