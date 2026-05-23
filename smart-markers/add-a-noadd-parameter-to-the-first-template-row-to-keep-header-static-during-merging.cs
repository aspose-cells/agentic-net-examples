using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Prepare a DataTable with a header row and some data rows
        DataTable table = new DataTable();
        table.Columns.Add("Header");
        table.Columns.Add("Value");
        table.Rows.Add("Header1", "Value1");
        table.Rows.Add("Header2", "Value2");
        table.Rows.Add("Header3", "Value3");

        // Import the DataTable into the worksheet.
        // ShiftFirstRowDown = false (no‑add) keeps the header in the first row
        // during any subsequent row insertions or merges.
        ImportTableOptions importOptions = new ImportTableOptions
        {
            ShiftFirstRowDown = false,
            IsFieldNameShown = true   // ensure the header row is imported
        };
        cells.ImportData(table, 0, 0, importOptions);

        // Example merge: merge the "Value" cells of the data rows (rows 2‑4, column B)
        cells.Merge(1, 1, 3, 1); // merges B2:B4 into a single cell

        // Keep the header row static when printing or paging by repeating it on each page
        worksheet.PageSetup.PrintTitleRows = "$1:$1";

        // Save the workbook
        workbook.Save("NoAddHeaderStatic.xlsx");
    }
}