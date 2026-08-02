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
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Existing data that will be present before the import operation
            // ------------------------------------------------------------
            // Populate some rows above the import area
            cells["A1"].PutValue("Header Above Import");
            cells["A2"].PutValue(123);
            cells["A3"].PutValue(456);
            cells["A4"].PutValue(789);
            cells["A5"].PutValue("Row before import start");

            // ------------------------------------------------------------
            // Prepare the DataTable to be imported
            // ------------------------------------------------------------
            DataTable table = new DataTable("Products");
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Price", typeof(decimal));

            table.Rows.Add(1, "Laptop", 999.99m);
            table.Rows.Add(2, "Smartphone", 699.49m);
            table.Rows.Add(3, "Tablet", 399.00m);

            // ------------------------------------------------------------
            // Configure import options:
            // - InsertRows = true  => new rows are added, existing rows are shifted down
            // - IsFieldNameShown = true (optional, shows column headers)
            // ------------------------------------------------------------
            ImportTableOptions options = new ImportTableOptions
            {
                InsertRows = true,
                IsFieldNameShown = true
            };

            // ------------------------------------------------------------
            // Import the DataTable starting at row index 5 (i.e., sixth row, zero‑based)
            // Existing rows from row 5 onward will be pushed down.
            // ------------------------------------------------------------
            int firstRowOffset = 5; // zero‑based index for the 6th row (A6)
            int firstColumn = 0;    // start at column A

            cells.ImportData(table, firstRowOffset, firstColumn, options);

            // ------------------------------------------------------------
            // Save the workbook to verify the result
            // ------------------------------------------------------------
            workbook.Save("ImportWithOffsetAndInsertRows.xlsx");
        }
    }
}