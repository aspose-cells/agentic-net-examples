using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a sample DataTable with column names and data
        DataTable dt = new DataTable("Sample");
        dt.Columns.Add("Product", typeof(string));
        dt.Columns.Add("Quantity", typeof(int));
        dt.Columns.Add("Price", typeof(double));

        dt.Rows.Add("Apple", 10, 0.5);
        dt.Rows.Add("Banana", 20, 0.3);
        dt.Rows.Add("Cherry", 15, 1.2);

        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Import the DataTable into the worksheet starting at cell A1.
        // Set IsFieldNameShown = true so column names become table headers.
        ImportTableOptions importOptions = new ImportTableOptions
        {
            IsFieldNameShown = true
        };
        ws.Cells.ImportData(dt, 0, 0, importOptions);

        // Calculate the range that contains the imported data (including header row)
        int totalRows = dt.Rows.Count + 1; // +1 for the header row
        int totalCols = dt.Columns.Count;

        // Add an Excel table (ListObject) over the imported range.
        // The last argument indicates that the first row is a header.
        ws.ListObjects.Add(0, 0, totalRows, totalCols, true);

        // Save the workbook to a file.
        wb.Save("DataTableToTable.xlsx");
    }
}