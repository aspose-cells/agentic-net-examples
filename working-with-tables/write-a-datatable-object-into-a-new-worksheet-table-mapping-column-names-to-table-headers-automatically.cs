using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsDataTableToTableDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add a worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Prepare a sample DataTable
            DataTable dt = new DataTable("Sample");
            dt.Columns.Add("Product", typeof(string));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("Price", typeof(double));

            dt.Rows.Add("Apple", 10, 0.5);
            dt.Rows.Add("Banana", 20, 0.3);
            dt.Rows.Add("Cherry", 15, 0.8);

            // Set import options to show column names as headers
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true   // map column names to table headers
            };

            // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
            int rowsImported = cells.ImportData(dt, 0, 0, importOptions);

            // Create an Excel table (ListObject) that covers the imported range
            // The range includes the header row plus the imported data rows
            int totalRows = rowsImported + 1; // +1 for header row
            int totalColumns = dt.Columns.Count;
            sheet.ListObjects.Add(0, 0, totalRows - 1, totalColumns - 1, true);

            // Save the workbook
            workbook.Save("DataTableToExcelTable.xlsx");
        }
    }
}