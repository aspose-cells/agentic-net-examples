using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsImportExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Prepare a sample DataTable to import
            DataTable table = new DataTable();
            table.Columns.Add("Product");
            table.Columns.Add("Quantity", typeof(int));
            table.Rows.Add("Apple", 10);
            table.Rows.Add("Banana", 20);
            table.Rows.Add("Cherry", 30);

            // Configure import options:
            // ShiftFirstRowDown = false ensures the first row is written at the specified position
            // (row index 0) without shifting existing rows down, thus overwriting any data present.
            ImportTableOptions options = new ImportTableOptions
            {
                ShiftFirstRowDown = false,
                IsFieldNameShown = false // Do not import column headers
            };

            // Import the data starting at row 0, column 0 (cell A1)
            cells.ImportData(table, 0, 0, options);

            // Save the workbook to a file
            workbook.Save("OverwriteImportDemo.xlsx");
        }
    }
}