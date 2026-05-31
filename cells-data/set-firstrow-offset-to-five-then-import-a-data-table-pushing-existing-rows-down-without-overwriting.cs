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
            // Add some existing data to demonstrate that it will be shifted
            // ------------------------------------------------------------
            // Existing data occupies rows 0 to 4 (A1 to A5)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue($"Existing Row {i + 1}");
            }

            // ------------------------------------------------------------
            // Prepare a DataTable that will be imported starting at row index 5 (i.e., row 6 in Excel)
            // ------------------------------------------------------------
            DataTable table = new DataTable("Products");
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Price", typeof(decimal));

            table.Rows.Add(101, "Laptop", 999.99m);
            table.Rows.Add(102, "Smartphone", 699.49m);
            table.Rows.Add(103, "Tablet", 399.00m);

            // ------------------------------------------------------------
            // Configure import options:
            //   InsertRows = true  -> new rows are inserted, pushing any existing rows down
            //   IsFieldNameShown = true -> column headers will be imported as the first row
            // ------------------------------------------------------------
            ImportTableOptions importOptions = new ImportTableOptions
            {
                InsertRows = true,
                IsFieldNameShown = true
            };

            // Import the DataTable starting at row index 5 (zero‑based), column 0 (A column)
            // Existing rows at and below this position will be shifted down.
            cells.ImportData(table, 5, 0, importOptions);

            // ------------------------------------------------------------
            // Save the workbook to verify the result
            // ------------------------------------------------------------
            workbook.Save("ImportWithOffsetAndInsertRows.xlsx");
        }
    }
}