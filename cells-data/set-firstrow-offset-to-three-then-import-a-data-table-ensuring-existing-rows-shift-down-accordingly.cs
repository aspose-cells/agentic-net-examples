using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsOffsetImportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add some existing data that we want to shift down
            cells["A1"].PutValue("Existing Row 1");
            cells["A2"].PutValue("Existing Row 2");
            cells["A3"].PutValue("Existing Row 3");

            // Insert three rows at the top (row index 0) to create an offset.
            // The existing rows will be moved down accordingly.
            cells.InsertRows(0, 3, true);

            // Prepare a sample DataTable to import
            DataTable table = new DataTable("Sample");
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Quantity", typeof(int));

            table.Rows.Add(101, "Apple", 50);
            table.Rows.Add(102, "Banana", 30);
            table.Rows.Add(103, "Cherry", 20);

            // Set import options (show column headers, do not insert additional rows)
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,
                InsertRows = false,
                ShiftFirstRowDown = false
            };

            // Import the DataTable starting at the first row (index 0) after the offset.
            cells.ImportData(table, 0, 0, importOptions);

            // Save the workbook
            workbook.Save("OffsetImportDemo.xlsx");
        }
    }
}