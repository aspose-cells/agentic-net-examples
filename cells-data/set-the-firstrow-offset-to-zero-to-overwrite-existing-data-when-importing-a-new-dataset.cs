using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsImportOverwriteDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Prepare a sample DataTable to import
            DataTable table = new DataTable("Sample");
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Quantity", typeof(int));

            table.Rows.Add(1, "Apple", 50);
            table.Rows.Add(2, "Banana", 30);
            table.Rows.Add(3, "Cherry", 20);

            // Configure import options:
            // - ShiftFirstRowDown = false ensures the first row is not shifted down.
            // - InsertRows = false makes the import overwrite existing cells instead of inserting new rows.
            // - IsFieldNameShown = false (optional) prevents column headers from being added.
            ImportTableOptions importOptions = new ImportTableOptions
            {
                ShiftFirstRowDown = false,
                InsertRows = false,
                IsFieldNameShown = false
            };

            // Import the data starting at row 0, column 0 (firstRow offset = 0)
            cells.ImportData(table, 0, 0, importOptions);

            // Save the workbook (save rule)
            workbook.Save("OverwriteImportDemo.xlsx");
        }
    }
}