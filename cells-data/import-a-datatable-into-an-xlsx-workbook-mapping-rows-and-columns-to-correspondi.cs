using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Build a sample DataTable with three columns
        DataTable dataTable = new DataTable("Products");
        dataTable.Columns.Add("ID", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Price", typeof(decimal));

        // Add some rows to the DataTable
        dataTable.Rows.Add(1, "Apple", 0.5m);
        dataTable.Rows.Add(2, "Banana", 0.3m);
        dataTable.Rows.Add(3, "Cherry", 1.2m);

        // Configure import options to include column headers in the worksheet
        ImportTableOptions importOptions = new ImportTableOptions();
        importOptions.IsFieldNameShown = true;

        // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
        cells.ImportData(dataTable, 0, 0, importOptions);

        // Save the workbook to an XLSX file
        workbook.Save("DataTableImportDemo.xlsx");
    }
}